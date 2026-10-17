//! Built-in step-debugger REPL — the client half of the debug surface. The VM
//! (in `vybe_runtime`) owns the pause/breakpoint state machine; this module is
//! pure transport + presentation: it attaches channels to the VM, spawns a
//! stdin-reader thread and an event-printer thread, and formats the typed
//! protocol for a terminal.
//!
//! The VM stays on the main thread (it is not `Send`); these worker threads hold
//! only channel endpoints, exactly like the browser debug server pattern.

use std::io::{BufRead, IsTerminal, Write};
use std::sync::mpsc::{Sender, channel};
use std::thread;

use vybe_runtime::debugger::{
    ChunkRef, DebugEvent, DebugResponse, FrameInfo, Location, PauseReason,
};
use vybe_runtime::{DebugCommand, DebugRequest, VM};

#[path = "debug_keyboard.rs"]
mod keyboard;

#[path = "debug_pointer.rs"]
mod pointer;

#[path = "debug_timing.rs"]
mod timing;

/// Created before CLI locals, so its destructor measures their teardown too.
pub(crate) struct CliReport {
    pub started: std::time::Instant,
    pub report: Option<vybe_runtime::debugger::SharedDebugReport>,
}
impl CliReport {
    pub fn new() -> Self {
        Self { started: std::time::Instant::now(), report: None ,}
    }
}
impl Drop for CliReport {
    fn drop(&mut self) {
        if let Some(report) = &self.report {
            let _ = std::io::stdout().flush();
            report.lock().unwrap().milestone("CLI teardown complete (before process exit)");
            print_report(report, None);
        }
    }
}

/// Render the debugger-owned timing snapshot, both on demand and after a
/// failed run, before the CLI exits and discards the collected phases.
pub(crate) fn print_report(report: &vybe_runtime::debugger::SharedDebugReport, filter: Option<&str>,) {
    let report = report.lock().unwrap();
    if report.process_started.is_some() {
        eprintln!("  CLI elapsed milestones (from entry; include debugger waits):");
        for (label, duration) in &report.milestones {
            eprintln!("    {:>8.3}s  {label}", duration.as_secs_f64());
        }
        if let Some(first) = report.first_output {
            eprintln!("    {:>8.3}s  first stdout write (may be buffered)", first.as_secs_f64());
        }
        if let Some(last) = report.last_output {
            eprintln!("    {:>8.3}s  last stdout write (may be buffered)", last.as_secs_f64());
        }
        eprintln!("    {:>8.3}s  debugger pause time", report.paused.as_secs_f64());
    }
    eprintln!("  instructions: {}  current: {}", report.instructions, report.location.as_deref().unwrap_or("before execution"));
    use std::sync::atomic::Ordering;
    let live = &report.live;
    eprintln!("  live: {}  chunk #{}@{} {:?}",
        live.instructions.load(Ordering::Relaxed),
        live.chunk.load(Ordering::Relaxed),
        live.ip.load(Ordering::Relaxed),
        vybe_runtime::opcode::Op(live.op.load(Ordering::Relaxed)));
    if let Some((name, start)) = &report.active_host {
        eprintln!("  host   {:>8.3}s  {name}", start.elapsed().as_secs_f64());
        if let Some(args) = &report.active_host_args {
            eprintln!("  host args: [{}]", args.join(", "));
        }
    }
    for (label, start) in &report.active {
        eprintln!("  active {:>8.3}s  {label}", start.elapsed().as_secs_f64());
    }
    let mut completed: Vec<_> = report.completed.iter()
        .filter(|(label, _)| filter.is_none_or(|text| label.contains(text)))
        .cloned().collect();
    let totals = ["compile ", "compiler total", "prepare source", "lower module", "vm link", "vm validate", "vm relocate", "vm globals", "vm decode", "vm blocks", "vm imports", "vm callsites", "vm types", "vm init globals",];
    for prefix in totals {
        let matches: Vec<_> = completed.iter().filter(|(label, _)| {
            if prefix == "compile " { label.starts_with(prefix) }
            else { label == prefix }
        }).collect();
        if !matches.is_empty() {
            let seconds: f64 = matches.iter().map(|(_, duration)| duration.as_secs_f64()).sum();
            eprintln!("  total {:>8.3}s  {:>5}  {prefix}", seconds, matches.len());
        }
    }
    let mut phases = std::collections::HashMap::<&str, (usize, f64)>::new();
    for (label, duration) in &completed {
        if label.starts_with("include ") || label.starts_with("compile ") || label.starts_with("function ") {
            continue;
        }
        let total = phases.entry(label.as_str()).or_default();
        total.0 += 1;
        total.1 += duration.as_secs_f64();
    }
    let mut phases: Vec<_> = phases.into_iter().collect();
    phases.sort_by(|a, b| b.1.1.total_cmp(&a.1.1));
    eprintln!("  grouped phases (nested timings overlap):");
    for (label, (count, seconds)) in phases.into_iter().take(20) {
        eprintln!("  total {seconds:>8.3}s  {count:>5}  {label}");
    }
    let mut hot: Vec<_> = report.function_samples.iter().collect();
    hot.sort_by(|a, b| b.1.cmp(a.1));
    if !hot.is_empty() {
        eprintln!("  sampled functions (1 sample / 4096 VM instructions):");
        for (name, samples) in hot.into_iter().take(12) {
            eprintln!("  {:>8}  {name}", samples);
        }
    }
    let mut hosts: Vec<_> = report.host_samples.iter().collect();
    hosts.sort_by(|a, b| b.1.1.cmp(&a.1.1));
    if !hosts.is_empty() {
        eprintln!("  sampled host wall time (1 call in 64; observed, nested work included):");
        for (name, (samples, elapsed)) in hosts.into_iter().take(12) {
            eprintln!("  {:>8.3}s  {:>6} samples  {name}", elapsed.as_secs_f64(), samples);
        }
    }
    let mut lines: Vec<_> = report.line_samples.iter().collect();
    lines.sort_by(|a, b| b.1.cmp(a.1));
    if !lines.is_empty() {
        eprintln!("  sampled source lines:");
        for (line, samples) in lines.into_iter().take(12) {
            eprintln!("  {:>8}  {line}", samples);
        }
    }
    completed.sort_by(|a, b| b.1.cmp(&a.1));
    eprintln!("  completed: {} phases", completed.len());
    for (label, elapsed) in completed.iter().take(12) {
        eprintln!("  {:>8.3}s  {label}", elapsed.as_secs_f64());
    }
}

fn reset_samples(report: &vybe_runtime::debugger::SharedDebugReport) -> u64 {
    let mut report = report.lock().unwrap();
    report.function_samples.clear();
    report.host_samples.clear();
    report.line_samples.clear();
    report.live.instructions.load(std::sync::atomic::Ordering::Relaxed)
}

/// Attach a debugger to `vm` and spawn the REPL worker threads. Call this before
/// running the VM; it pauses on entry so breakpoints can be set first.
pub fn attach(vm: &mut VM) {
    let (cmd_tx, cmd_rx) = channel::<DebugRequest>();
    let (evt_tx, evt_rx) = channel::<DebugEvent>();
    vm.attach_debugger(cmd_rx, evt_tx, /* pause_on_entry */ true);
    let report = vm.debug_report().expect("attached debugger has a report");
    // Runs on the VM's thread, before the guest does — so the document opened
    // here is the one the guest will build in, and `widgets` can read it from
    // the REPL thread.
    crate::gui_document::pin();

    // Event printer: renders async events (paused / resumed / exited / opcode).
    thread::spawn(move || {
        let mut timer = timing::RunTimer::default();
        for event in evt_rx {
            let now = std::time::Instant::now();
            match &event {
                DebugEvent::Resumed => timer.resume(now),
                DebugEvent::Paused { .. } | DebugEvent::Exited { .. } => {
                    if let Some(elapsed) = timer.stop(now) {
                        eprintln!("  run interval (observer wall): {:.6}s", elapsed.as_secs_f64());
                    }
                }
                _ => {},
            }
            print_event(&event);
        }
    });

    // Stdin reader: parse a line → command → send → print the reply.
    thread::spawn(move || {
        let stdin = std::io::stdin();
        let interactive = stdin.is_terminal();
        let mut lines = stdin.lock().lines();
        banner();
        loop {
            prompt();
            let Some(Ok(line)) = lines.next() else { break };
            let line = line.trim();
            if line.is_empty() {
                continue;
            }
            // `widgets`/`controls` are handled entirely client-side: they walk
            // the live document, which is safe whether the VM is paused or
            // running, so they never round-trip through the VM.
            let head = line.split_whitespace().next().unwrap_or("");
            if head == "report" {
                let filter = line.split_once(' ').map(|(_, text)| text.trim()).filter(|text| !text.is_empty());
                if filter == Some("reset-samples") {
                    let instructions = reset_samples(&report);
                    eprintln!("  samples cleared at {instructions} live instructions; timings and execution state preserved");
                } else {
                    print_report(&report, filter);
                }
                continue;
            }
            if matches!(head, "widgets" | "controls") {
                print_document_widgets();
                continue;
            }
            // `capture` is client-side for the same reason: it renders the live
            // document into an offscreen pixmap.
            if head == "capture" {
                capture_frame(&line.split_whitespace().skip(1).collect::<Vec<_>>());
                continue;
            }
            if head == "pointer" {
                match pointer::parse(&line.split_whitespace().skip(1).collect::<Vec<_>>()) {
                    Ok(event) => {
                        use vybe_platform_web::engine::{self, DomOp};
                        engine::apply(crate::gui_document::active(), DomOp::DispatchPointer {
                            kind: event.kind, client_x: event.x, client_y: event.y, button: event.button,
                        });
                        eprintln!("  pointer event dispatched through browser hit-testing");
                    }
                    Err(error) => eprintln!("  {error}"),
                }
                continue;
            }
            if head == "key" {
                match keyboard::parse(&line.split_whitespace().skip(1).collect::<Vec<_>>()) {
                    Ok(event) => {
                        use vybe_platform_web::engine::{self, DomOp, UiEventFields};
                        engine::apply(crate::gui_document::active(), DomOp::DispatchKeyboard(UiEventFields {
                            kind: event.kind, key: event.key, key_code: event.key_code, code: event.code,
                            ..Default::default()
                        }));
                        eprintln!("  keyboard event dispatched; guest handlers run through the browser event queue");
                    }
                    Err(error) => eprintln!("  {error}"),
                }
                continue;
            }
            // `html` is client-side for the same reason as `widgets`: it reads
            // the live document, which is safe to walk whether the VM is
            // paused or running.
            if matches!(head, "html" | "dom") {
                print_html(&line.split_whitespace().skip(1).collect::<Vec<_>>());
                continue;
            }
            // The inspector: read or write one element, live. Client-side for
            // the same reason as `widgets` — it walks the document, which is
            // safe whether the VM is paused or running.
            if matches!(head, "css" | "attr" | "text") {
                inspect(head, &line.split_whitespace().skip(1).collect::<Vec<_>>());
                continue;
            }
            if matches!(head, "draws" | "drawlist") {
                print_draws(&line.split_whitespace().skip(1).collect::<Vec<_>>());
                continue;
            }
            // `trace canvas on|off` is client-side — it flips a process-wide
            // toggle the host draw path reads, so it needs no VM round-trip and
            // works while the VM is running.
            if head == "trace" {
                let rest: Vec<&str> = line.split_whitespace().skip(1).collect();
                if rest.first() == Some(&"canvas") {
                    let on = rest.get(1) != Some(&"off");
                    webcore::canvas::set_trace_enabled(on);
                    eprintln!("  canvas tracing {}", if on { "on" } else { "off" });
                    continue;
                }
            }
            match parse_command(line) {
                Ok(command) => {
                    // Interactive GUI sessions must accept client-side commands
                    // while running. A piped session instead needs `continue`
                    // to wait for the next stop before reading another command;
                    // otherwise `bt`/`locals` race ahead into the running VM.
                    let async_continue =
                        interactive && matches!(command, DebugCommand::Continue | DebugCommand::Pause);
                    if async_continue {
                        if cmd_tx
                            .send(DebugRequest {
                                command,
                                reply: channel::<DebugResponse>().0,
                            })
                            .is_err()
                        {
                            break; // channel closed — VM gone
                        }
                    } else if !send_and_print(&cmd_tx, command) {
                        break;
                    }
                }
                Err(msg) => eprintln!("  {msg}"),
            }
        }
    });
}

/// `draws [control]` — what is actually ON each canvas.
///
/// This is what tells "nothing was drawn" apart from "drawn in the wrong
/// place" — two failures that look identical on screen.
///
/// It used to list the recorded draw COMMANDS. A canvas paints into its bitmap
/// as the calls arrive now, so there is no command list to print; there are
/// pixels, which answer the same two questions more directly — how much ink
/// landed, and where it landed. The one thing the old form could tell you and
/// this cannot is "drawn, then painted over", because a bitmap keeps only the
/// result. Reported honestly below rather than left to be inferred.
fn print_draws(args: &[&str]) {
    let wanted = args
        .iter()
        .find(|a| a.parse::<usize>().is_err())
        .map(|w| w.to_lowercase());

    /// Size, ink, and the bounding box of the ink.
    struct Ink {
        w: u32,
        h: u32,
        inked: usize,
        bounds: Option<(u32, u32, u32, u32)>,
    }

    // **A drawing surface is a `<canvas>` ELEMENT in the document.**
    let mut found: Vec<(String, Ink)> = Vec::new();
    use vybe_platform_web::canvas_backend::{self, Query2D, Query2DValue};
    use vybe_platform_web::engine::{self, DomOp, DomValue};
    let document = crate::gui_document::active();
    if let DomValue::Nodes(nodes) = engine::apply(document, DomOp::ElementsByTag("canvas".into())) {
        for node in nodes {
            // Report the name a caller would recognise — the `id`/`name` they
            // gave it — falling back to the internal node name.
            let attribute = |name: &str| match engine::apply(
                document,
                DomOp::GetAttribute(node, name.into()),
            ) {
                DomValue::Text(value) if !value.is_empty() => Some(value),
                _ => None,
            };
            let label = attribute("id").or_else(|| attribute("name"))
                .unwrap_or_else(|| format!("n{node}"));
            let (w, h) = match engine::apply(document, DomOp::CanvasSize(node)) {
                DomValue::Pair(w, h) => (w as u32, h as u32),
                _ => continue,
            };
            let target = format!("d{document}:n{node}");
            let (Ok(sw), Ok(sh)) = (i32::try_from(w), i32::try_from(h)) else {
                continue;
            };
            let Query2DValue::Pixels {
                data,
                width: w,
                height: h,
            } = canvas_backend::query(
                &target,
                Query2D::GetImageData {
                    sx: 0,
                    sy: 0,
                    sw,
                    sh,
                },
            )
            else {
                continue;
            };
            let ink = {
                let mut inked = 0usize;
                let (mut x0, mut y0, mut x1, mut y1) = (u32::MAX, u32::MAX, 0u32, 0u32);
                for (i, px) in data.chunks_exact(4).enumerate() {
                    if px[3] == 0 {
                        continue;
                    }
                    inked += 1;
                    let (x, y) = (i as u32 % w, i as u32 / w);
                    x0 = x0.min(x);
                    y0 = y0.min(y);
                    x1 = x1.max(x);
                    y1 = y1.max(y);
                }
                Ink {
                    w,
                    h,
                    inked,
                    bounds: (inked > 0).then_some((x0, y0, x1, y1)),
                }
            };
            found.push((label, ink));
        }
    }

    if found.is_empty() {
        eprintln!("  (no canvases exist)");
        return;
    }

    let mut shown = false;
    for (name, ink) in &found {
        if let Some(w) = &wanted {
            // Same forgiving match as `--capture-control`: a canvas named after
            // a window title is not something anyone types exactly.
            if !name.to_lowercase().contains(w.as_str()) {
                continue;
            }
        }
        shown = true;
        match ink.bounds {
            None => eprintln!(
                "  `{name}`  {}x{}  NOTHING DRAWN (every pixel transparent)",
                ink.w, ink.h
            ),
            Some((x0, y0, x1, y1)) => {
                let total = (ink.w as usize) * (ink.h as usize);
                let pct = 100.0 * ink.inked as f32 / total.max(1) as f32;
                eprintln!(
                    "  `{name}`  {}x{}  {} px inked ({pct:.1}%)  ink bounds {x0},{y0} → {x1},{y1}",
                    ink.w, ink.h, ink.inked
                );
                if x1 >= ink.w || y1 >= ink.h {
                    eprintln!("        ⚠ ink reaches the bitmap edge — it may be clipped");
                }
            }
        }
    }
    if !shown {
        let names: Vec<&str> = found.iter().map(|(n, _)| n.as_str()).collect();
        eprintln!(
            "  no canvas named `{}` (have: {})",
            wanted.unwrap_or_default(),
            names.join(", ")
        );
    }
}

/// `capture [control] [file]` — write the live frame to a PNG.
///
/// Both arguments are optional: with none it writes the whole form to
/// `vybe-capture.png`. A single argument ending in `.png` is taken as the file,
/// otherwise as a control name.
fn capture_frame(args: &[&str]) {
    let is_file = |s: &str| s.ends_with(".png");
    let (control, path) = match args {
        [] => (None, "vybe-capture.png"),
        [one] if is_file(one) => (None, *one),
        [one] => (Some(*one), "vybe-capture.png"),
        [a, b, ..] => (Some(*a), *b),
    };
    match crate::gui_capture::capture_to_png(path, control, 1.0) {
        Ok((w, h)) => eprintln!("  wrote {w}x{h} PNG → {path}"),
        Err(e) => eprintln!("  capture failed: {e}"),
    }
}

/// Dump the document's elements — geometry, observable properties, listeners.
///
/// A control IS `document.createElement(tag)` for every frontend, so the
/// document is what a running program has built.
/// `html` — the live document as markup.
///
/// The companion to `widgets`, not a replacement: that one reports each
/// element's properties and laid-out rect, this one reports the STRUCTURE.
/// A control with correct properties in the wrong parent reads as fine in a
/// property dump and is obvious here.
///
/// A designer form built straight into `GuiState` never opens a document, so
/// there is genuinely nothing to serialise — say so rather than print an empty
/// body and imply the program built nothing.
/// Do two CSS lengths mean the same thing? `8` and `8px` do — the emitters
/// still write geometry unitless, so a bare number and a `px` value are the
/// same declaration and flagging them as a mismatch would cry wolf on every
/// control.
fn same_length(a: &str, b: &str) -> bool {
    let strip = |s: &str| s.trim().trim_end_matches("px").trim().to_string();
    strip(a) == strip(b)
}

fn print_html(args: &[&str]) {
    match args.first() {
        // One element and its subtree — `outerHTML`. Useful on a form with
        // sixty controls, where the whole body scrolls past.
        Some(name) => match crate::gui_document::node_by_id(name) {
            Some(node) => match crate::gui_document::inspect::outer_html(node) {
                Some(html) => println!("{html}"),
                None => eprintln!("  (no document)"),
            },
            None => eprintln!("  no control named `{name}` (see `widgets` for names)"),
        },
        None => match crate::gui_document::html() {
            Some(html) => println!("{html}"),
            None => eprintln!("  (no document — nothing has built one yet)"),
        },
    }
}

/// The inspector: `css` / `attr` / `text`, each reading with one argument and
/// writing with two.
///
/// Writes go through the same `Document` entry points the guest uses, so
/// setting something here does exactly what the program setting it would do —
/// and a change is visible in `widgets`, `html` and `capture` immediately,
/// because there is only one tree.
fn inspect(verb: &str, args: &[&str]) {
    let usage = match verb {
        "css" => "usage: css <control> [property] [value]",
        "attr" => "usage: attr <control> <name> [value]",
        _ => "usage: text <control> [value]",
    };
    let Some(name) = args.first() else {
        eprintln!("  {usage}");
        return;
    };
    let Some(node) = crate::gui_document::node_by_id(name) else {
        eprintln!("  no control named `{name}` (see `widgets` for names)");
        return;
    };
    use crate::gui_document::inspect;
    match (verb, args.get(1), args.get(2)) {
        // `css <control>` — every declaration on the element.
        ("css", None, _) => match inspect::declarations(node) {
            Some(declarations) if !declarations.is_empty() => {
                for (property, value) in declarations {
                    println!("  {property}: {value}");
                }
            }
            Some(_) => eprintln!("  (no declarations)"),
            None => eprintln!("  (no document)"),
        },
        // Declared vs computed, and SAY when they differ. Geometry is read off
        // the control, so a laid-out or docked element does not sit where its
        // own `left` asked — and that gap is the single most useful number in a
        // layout bug. Printing one figure hides it.
        ("css", Some(property), None) => {
            let computed = inspect::style(node, property);
            let declared = inspect::declarations(node).and_then(|declarations| {
                declarations
                    .into_iter()
                    .find(|(name, _)| name.eq_ignore_ascii_case(property))
                    .map(|(_, value)| value)
            });
            match (declared, computed) {
                (None, None) => eprintln!("  (no document)"),
                (None, Some(value)) if value.is_empty() => {
                    eprintln!("  {property}: (not set)")
                }
                (None, Some(value)) => println!("  {property}: {value}  (computed)"),
                (Some(declared), Some(computed))
                    if !computed.is_empty() && !same_length(&declared, &computed) =>
                {
                    println!("  {property}: {declared}  ← declared, computed {computed}");
                }
                (Some(declared), _) => println!("  {property}: {declared}"),
            }
        }
        ("css", Some(property), Some(_)) => {
            // Everything after the property is the value, so
            // `css btn font "16px Helvetica"` works unquoted too.
            let value = args[2..].join(" ");
            inspect::set_style(node, property, &value);
            println!("  {property}: {value}");
        }
        ("attr", Some(attribute), None) => match inspect::attribute(node, attribute) {
            Some(value) => println!("  {attribute}=\"{value}\""),
            None => eprintln!("  {attribute}: (not set)"),
        },
        ("attr", Some(attribute), Some(_)) => {
            let value = args[2..].join(" ");
            inspect::set_attribute(node, attribute, &value);
            println!("  {attribute}=\"{value}\"");
        }
        ("text", None, _) => match inspect::text(node) {
            Some(text) => println!("  {text:?}"),
            None => eprintln!("  (no document)"),
        },
        ("text", Some(_), _) => {
            let value = args[1..].join(" ");
            inspect::set_text(node, &value);
            println!("  {value:?}");
        }
        _ => eprintln!("  {usage}"),
    }
}

fn print_document_widgets() -> bool {
    let controls = crate::gui_document::controls();
    if controls.is_empty() {
        return false;
    }
    match crate::gui_document::viewport() {
        Some((w, h)) => eprintln!("  document {w}×{h}  ({} element(s))", controls.len()),
        None => eprintln!("  document  ({} element(s))", controls.len()),
    }
    for control in controls {
        // A control with no rect will not render and cannot be hit-tested —
        // but SAY WHICH of the two reasons it is. `rect` is looked up in the
        // form's tree, so a missing one usually means the element was never
        // appended, not that layout has not run; those are a missing
        // `appendChild` and a missing layout pass, and reporting them the same
        // way sent a session hunting container layout for a control that was
        // never in the container.
        let rect = match (control.rect, control.connected) {
            (Some(r), _) if r.w >= 1.0 && r.h >= 1.0 => {
                format!("  rect={},{} {}x{}", r.x, r.y, r.w, r.h)
            }
            (Some(_), _) => "  rect=0x0 ← never laid out".to_string(),
            (None, false) => "  ⚠ DETACHED ← created, never appended".to_string(),
            (None, true) => "  rect=? ← in the document, not laid out".to_string(),
        };
        let props = if control.properties.is_empty() {
            String::new()
        } else {
            let pairs: Vec<String> = control
                .properties
                .iter()
                .map(|(k, v)| format!("{k}={v}"))
                .collect();
            format!("  {{{}}}", pairs.join(", "))
        };
        let events = if control.events.is_empty() {
            String::new()
        } else {
            format!("  events[{}]", control.events.join(","))
        };
        // An unnamed control still needs a handle a `click` can take, and
        // `n<node>` is the one the document itself uses.
        let handle = if control.id.is_empty() {
            format!("n{}", control.node)
        } else {
            control.id.clone()
        };
        eprintln!("  • {handle} <{}>{rect}{props}{events}", control.tag);
    }
    true
}

fn banner() {
    eprintln!("── vybe step debugger ── type `h` for help. Paused on entry.");
}

fn prompt() {
    eprint!("(vdbg) ");
    let _ = std::io::stderr().flush();
}

/// Send a command, block on its reply, print it. Returns false if the VM channel
/// is gone.
fn send_and_print(cmd_tx: &Sender<DebugRequest>, command: DebugCommand) -> bool {
    let (reply_tx, reply_rx) = channel::<DebugResponse>();
    if cmd_tx
        .send(DebugRequest {
            command,
            reply: reply_tx,
        })
        .is_err()
    {
        return false;
    }
    match reply_rx.recv() {
        Ok(resp) => {
            print_response(&resp);
            true
        }
        Err(_) => false,
    }
}

// ─── Command parsing ────────────────────────────────────────────────────────

fn parse_memory_arguments(parts: &[&str]) -> Result<(u32, u64, u32), String> {
    if !(2..=3).contains(&parts.len()) {
        return Err("usage: memory <index> <address> [byte-count]".into());
    }
    let number = |text: &str| {
        if let Some(hex) = text.strip_prefix("0x").or_else(|| text.strip_prefix("0X")) {
            u64::from_str_radix(hex, 16)
        } else {
            text.parse::<u64>()
        }
    };
    let memory = u32::try_from(number(parts[0]).map_err(|_| "bad memory index")?)
        .map_err(|_| "memory index exceeds u32")?;
    let address = number(parts[1]).map_err(|_| "bad memory address")?;
    let length = match parts.get(2) {
        Some(text) => u32::try_from(number(text).map_err(|_| "bad byte count")?)
            .map_err(|_| "byte count exceeds u32")?,
        None => 64,
    };
    Ok((memory, address, length))
}

fn memory_dump_rows(address: u64, bytes: &[u8]) -> Vec<String> {
    bytes.chunks(16).enumerate().map(|(row, chunk)| {
        let hex = chunk.iter().map(|byte| format!("{byte:02x}")).collect::<Vec<_>>().join(" ");
        let text: String = chunk.iter().map(|&byte| {
            if byte.is_ascii_graphic() || byte == b' ' { char::from(byte) } else { '.' }
        }).collect();
        let offset = u128::from(address) + row as u128 * 16;
        format!("{offset:016x}  {hex:47}  |{text}|")
    }).collect()
}

fn parse_command(line: &str) -> Result<DebugCommand, String> {
    let mut parts = line.split_whitespace();
    let cmd = parts.next().unwrap_or("");
    let rest: Vec<&str> = parts.collect();
    Ok(match cmd {
        "c" | "cont" | "continue" => DebugCommand::Continue,
        "pause" | "interrupt" => DebugCommand::Pause,
        "s" | "step" => DebugCommand::StepInto,
        "n" | "next" => DebugCommand::StepOver,
        "o" | "out" | "fin" | "finish" => DebugCommand::StepOut,
        "si" | "stepi" => DebugCommand::StepInstruction,
        "detach" => DebugCommand::Detach,
        "q" | "quit" | "exit" => DebugCommand::Quit,

        "b" | "break" => parse_breakpoint(&rest)?,
        "bl" | "breaks" => DebugCommand::ListBreakpoints,
        "bd" | "delete" => {
            let id = rest
                .first()
                .ok_or("usage: bd <id>")?
                .parse()
                .map_err(|_| "bad id")?;
            DebugCommand::DeleteBreakpoint { id }
        }
        "enable" => DebugCommand::EnableBreakpoint {
            id: rest
                .first()
                .ok_or("usage: enable <id>")?
                .parse()
                .map_err(|_| "bad id")?,
            enabled: true,
        },
        "disable" => DebugCommand::EnableBreakpoint {
            id: rest
                .first()
                .ok_or("usage: disable <id>")?
                .parse()
                .map_err(|_| "bad id")?,
            enabled: false,
        },

        "p" | "print" => {
            // Join the whole tail so compound expressions (`p a + b`) survive;
            // the debugger tries a fast structural read first, then full eval.
            let path = rest.join(" ");
            if path.is_empty() {
                return Err("usage: p <name>[.field][idx]  or  p <expr>".into());
            }
            DebugCommand::Print { path }
        }
        "memory" | "mem" => {
            let (memory, address, length) = parse_memory_arguments(&rest)?;
            DebugCommand::ReadMemory { memory, address, length }
        }
        "set" => {
            let joined = rest.join(" ");
            let (name, val) = joined
                .split_once('=')
                .ok_or("usage: set <name> = <literal>")?;
            DebugCommand::SetVar {
                name: name.trim().to_string(),
                literal: val.trim().to_string(),
            }
        }
        "bt" | "where" | "backtrace" => DebugCommand::Backtrace,
        "locals" | "l" => DebugCommand::Locals {
            frame: rest.first().and_then(|s| s.parse().ok()).unwrap_or(0),
        },
        "stack" => DebugCommand::OperandStack,
        "globals" | "g" => DebugCommand::Globals {
            prefix: rest.first().map(|s| s.to_string()),
        },
        "trace-global-inits" => DebugCommand::TraceGlobalInits {
            enabled: rest.first() != Some(&"off"),
        },
        "bgi" => DebugCommand::BreakGlobalInit {
            name: rest.first().ok_or("usage: bgi <global-name|*|off>")?.to_string(),
        },
        "dis" | "disasm" => {
            let offset = rest
                .first()
                .and_then(|s| s.strip_prefix('@'))
                .map(|s| {
                    s.parse::<usize>()
                        .map_err(|_| "usage: dis [@offset] [window]")
                })
                .transpose()?;
            let window = rest
                .get(usize::from(offset.is_some()))
                .map(|s| {
                    s.parse::<usize>()
                        .map_err(|_| "usage: dis [@offset] [window]")
                })
                .transpose()?
                .unwrap_or(4);
            DebugCommand::Disasm { window, offset }
        }
        "targets" => DebugCommand::ControlTargets {
            offset: rest.first().ok_or("usage: targets <offset>")?
                .parse().map_err(|_| "usage: targets <offset>")?,
        },
        "chunks" => DebugCommand::Chunks,
        "frame" | "fr" => DebugCommand::Frame {
            frame: rest.first().and_then(|s| s.parse().ok()).unwrap_or(0),
        },
        "watch" | "w" => {
            let expr = rest.join(" ");
            if expr.is_empty() {
                DebugCommand::ListWatches
            } else {
                DebugCommand::AddWatch { expr }
            }
        }
        "watches" => DebugCommand::ListWatches,
        "unwatch" | "clearwatch" | "clearwatches" => DebugCommand::ClearWatches,
        "reload" | "r" => DebugCommand::Reload,
        "restart" | "R" => DebugCommand::Restart,
        "trace" => DebugCommand::StreamOpcodes {
            enabled: rest.first() != Some(&"off"),
        },
        "skip-system" | "sys" => DebugCommand::SetSkipSystem {
            enabled: rest.first() != Some(&"off"),
        },

        // ── simulate GUI events (fire a handler through the live VM) ──
        "click" | "tap" => DebugCommand::FireEvent {
            control: rest
                .first()
                .ok_or("usage: click <control>  (see `widgets` for names)")?
                .to_string(),
            event: "click".to_string(),
        },
        "fire" => {
            let control = rest
                .first()
                .ok_or("usage: fire <control> <event>")?
                .to_string();
            // Folded HERE, not in the DOM: a person types `Click` and means
            // the `click` event, and that convenience belongs to the command
            // rather than to `addEventListener`, whose types are
            // case-SENSITIVE (DOM §2.7).
            let event = rest
                .get(1)
                .ok_or("usage: fire <control> <event>")?
                .to_ascii_lowercase();
            DebugCommand::FireEvent { control, event }
        }
        "close" | "window-close" => DebugCommand::FireEvent {
            control: rest.first().unwrap_or(&"form").to_string(),
            event: "Close".to_string(),
        },

        // ── function breakpoint / logpoint / run-to-cursor / ignore ──
        "bf" | "break-fn" => {
            let name = rest
                .first()
                .ok_or("usage: bf <function> [if <cond>]")?
                .to_string();
            let condition = rest
                .iter()
                .position(|t| *t == "if")
                .map(|i| rest[i + 1..].join(" "))
                .filter(|c| !c.trim().is_empty());
            DebugCommand::BreakFunction { name, condition }
        }
        "bh" | "break-host" => {
            let name = rest.first().ok_or("usage: bh <module>:<name> [invalid]|off")?;
            if rest.len() > 1 && !matches!(rest.get(1), Some(&"nonstring") | Some(&"invalid")) {
                return Err("usage: bh <module>:<name> [invalid]|off".into());
            }
            DebugCommand::BreakHost {
                name: (*name != "off").then(|| (*name).to_string()),
                bad_type_only: matches!(rest.get(1), Some(&"nonstring") | Some(&"invalid")),
            }
        }
        "logpoint" | "lp" => {
            let line = rest
                .first()
                .ok_or("usage: lp <line> <message>")?
                .parse()
                .map_err(|_| "bad line")?;
            let message = rest[1..].join(" ");
            if message.is_empty() {
                return Err("usage: lp <line> <message>  (use {expr} to interpolate)".into());
            }
            DebugCommand::Logpoint { line, message }
        }
        "runto" | "rt" | "tbreak" => {
            let line = rest
                .first()
                .ok_or("usage: rt <line>")?
                .parse()
                .map_err(|_| "bad line")?;
            DebugCommand::RunToLine { line }
        }
        "ignore" => {
            let id = rest
                .first()
                .ok_or("usage: ignore <id> <count>")?
                .parse()
                .map_err(|_| "bad id")?;
            let count = rest
                .get(1)
                .ok_or("usage: ignore <id> <count>")?
                .parse()
                .map_err(|_| "bad count")?;
            DebugCommand::SetIgnoreCount { id, count }
        }
        // ── exception breakpoints ──
        "catch" => match rest.first().copied() {
            Some("throw") => DebugCommand::ExceptionBreak {
                on_throw: true,
                on_uncaught: false,
            },
            Some("uncaught") => DebugCommand::ExceptionBreak {
                on_throw: false,
                on_uncaught: true,
            },
            Some("off") | None => DebugCommand::ExceptionBreak {
                on_throw: false,
                on_uncaught: false,
            },
            Some(other) => return Err(format!("usage: catch throw|uncaught|off  (got `{other}`)")),
        },
        // ── data watchpoints ──
        "wp" | "watchpoint" => {
            let target = rest.join(" ");
            if target.is_empty() {
                DebugCommand::ListWatchpoints
            } else {
                DebugCommand::AddWatchpoint { target }
            }
        }
        "wps" => DebugCommand::ListWatchpoints,
        "unwp" | "clearwp" => DebugCommand::ClearWatchpoints,
        // ── fibers / threads ──
        "fibers" | "threads" => DebugCommand::Fibers,

        "h" | "help" | "?" => {
            print_help();
            return Err(String::new());
        }
        other => return Err(format!("unknown command `{other}` — try `h`")),
    })
}

/// `b <chunk>:<line> [if <cond>]`, `b <chunk>@<offset> [if <cond>]`. Chunk may
/// be a name or an index. An `if <expr>` suffix makes it a conditional
/// breakpoint (evaluated in the paused frame).
fn parse_breakpoint(rest: &[&str]) -> Result<DebugCommand, String> {
    let spec = rest.first().ok_or("usage: b <chunk>:<line> [if <cond>]")?;
    // Optional `if <condition>` (everything after the `if` token).
    let condition = rest
        .iter()
        .position(|t| *t == "if")
        .map(|i| rest[i + 1..].join(" "))
        .filter(|c| !c.trim().is_empty());
    // Bare line number → break on that source line across all chunks.
    if let Ok(line) = spec.parse::<u32>() {
        return Ok(DebugCommand::BreakSourceLine { line, condition });
    }
    if let Some((chunk, line)) = spec.split_once(':') {
        let line: u32 = line.parse().map_err(|_| "bad line number")?;
        Ok(DebugCommand::BreakLine {
            chunk: chunk_ref(chunk),
            line,
            condition,
        })
    } else if let Some((chunk, offset)) = spec.split_once('@') {
        let offset: usize = offset.parse().map_err(|_| "bad offset")?;
        Ok(DebugCommand::BreakOffset {
            chunk: chunk_ref(chunk),
            offset,
            condition,
        })
    } else {
        Err("usage: b <line>  ·  b <file>:<line>  ·  b <chunk>@<offset>  [if <cond>]".into())
    }
}

fn chunk_ref(s: &str) -> ChunkRef {
    match s.parse::<usize>() {
        Ok(i) => ChunkRef::Index(i),
        Err(_) => ChunkRef::Name(s.to_string()),
    }
}

// ─── Presentation ───────────────────────────────────────────────────────────

fn print_event(event: &DebugEvent) {
    match event {
        DebugEvent::Paused {
            reason,
            location,
            frame_summary,
            watches,
        } => {
            eprintln!(
                "\n■ paused ({}) — {}",
                reason_str(reason),
                fmt_location(location)
            );
            eprintln!("  {frame_summary}");
            for (expr, val) in watches {
                eprintln!("  ◦ {expr} = {val}");
            }
            prompt();
        }
        DebugEvent::Resumed => {}
        DebugEvent::Exited { value } => {
            eprintln!("\n● program exited → {value}");
        }
        DebugEvent::Opcode {
            chunk,
            ip,
            op,
            stack_depth,
        } => {
            eprintln!("  · {chunk}@{ip:04} {op}  (stack {stack_depth})");
        }
        DebugEvent::Log { message } => eprintln!("  {message}"),
    }
}

/// Collapse a rendered value to ONE line, bounded.
///
/// A frame dump is read as a table. A value holding a newline — every buffered
/// stdout slot does — splits its own row and makes the column of slot indices
/// unreadable, which is the single thing that view exists to show.
fn one_line(value: &str) -> String {
    const MAX: usize = 72;
    let flat: String = value
        .chars()
        .map(|c| match c {
            '\n' => '␊',
            '\r' => '␍',
            '\t' => '→',
            c => c,
        })
        .collect();
    if flat.chars().count() > MAX {
        let head: String = flat.chars().take(MAX).collect();
        format!("{head}… ({} chars)", value.chars().count())
    } else {
        flat
    }
}

fn print_response(resp: &DebugResponse) {
    match resp {
        DebugResponse::Ok => {}
        DebugResponse::Error(e) => eprintln!("  error: {e}"),
        DebugResponse::Value(s) => eprintln!("  {s}"),
        DebugResponse::Memory { memory, address, bytes } => {
            eprintln!("  memory #{memory} @0x{address:x}: {} byte(s)", bytes.len());
            for row in memory_dump_rows(*address, bytes) {
                eprintln!("  {row}");
            }
        }
        DebugResponse::Fibers(lines) => {
            for l in lines {
                eprintln!("  {l}");
            }
        }
        DebugResponse::BreakpointSet {
            id,
            chunk,
            offset,
            line,
        } => {
            let at = line
                .map(|l| format!("line {l}"))
                .unwrap_or_else(|| format!("offset {offset}"));
            eprintln!("  breakpoint #{id} set at {chunk} {at}");
        }
        DebugResponse::Breakpoints(bps) => {
            if bps.is_empty() {
                eprintln!("  (no breakpoints)");
            }
            for b in bps {
                let en = if b.enabled { "" } else { " (disabled)" };
                let line = b.line.map(|l| format!(":{l}")).unwrap_or_default();
                eprintln!("  #{} {}@{}{}{}", b.id, b.chunk_name, b.offset, line, en);
            }
        }
        DebugResponse::Backtrace(frames) => {
            for f in frames {
                eprintln!("  {}", fmt_frame(f));
            }
        }
        DebugResponse::Locals(slots) => {
            if slots.is_empty() {
                eprintln!("  (no locals)");
            }
            for s in slots {
                match &s.name {
                    Some(name) => eprintln!("  {} [{}] = {}", name, s.index, s.value),
                    None => eprintln!("  [{}] = {}", s.index, s.value),
                }
            }
        }
        DebugResponse::Frame(shape) => {
            eprintln!(
                "  frame #{} — chunk [{}] {} (arity {})",
                shape.frame, shape.chunk_index, shape.chunk_name, shape.arity
            );
            // The layout facts first: a capture region overlapping the
            // parameters, or a local_count below the highest slot in use, is
            // the bug itself rather than a symptom of one.
            eprintln!(
                "  locals {} · scratch high-water {} · captures {}..{} ({}) · stack base {}",
                shape.local_count,
                shape.scratch_high_water,
                shape.capture_base,
                shape.capture_base as usize + shape.capture_count as usize,
                shape.capture_count,
                shape.base,
            );
            if shape.slots.is_empty() {
                eprintln!("  (frame declares no locals)");
            }
            for s in &shape.slots {
                let name = s.name.as_deref().unwrap_or("-");
                let dead = if s.live { "" } else { "  (beyond frame)" };
                eprintln!(
                    "  [{:>3}] {:<8} {:<28} = {} [{}]{}",
                    s.index,
                    s.owner.tag(),
                    name,
                    one_line(&s.value),
                    s.storage,
                    dead
                );
            }
        }
        DebugResponse::OperandStack(vals) => {
            if vals.is_empty() {
                eprintln!("  (operand stack empty)");
            }
            for (i, v) in vals.iter().enumerate() {
                eprintln!("  {i}: {v}");
            }
        }
        DebugResponse::Globals(pairs) => {
            if pairs.is_empty() {
                eprintln!("  (no matching globals)");
            }
            for (k, v) in pairs {
                eprintln!("  {k} = {v}");
            }
        }
        DebugResponse::Disasm { current_ip, lines } => {
            for l in lines {
                let marker = if l.is_current { "▶" } else { " " };
                let ln = l.line.map(|n| format!("  ; line {n}")).unwrap_or_default();
                eprintln!("  {marker} {:04}  {}{}", l.offset, l.text, ln);
            }
            let _ = current_ip;
        }
        DebugResponse::Chunks(chunks) => {
            for c in chunks {
                eprintln!(
                    "  [{}] {} (arity {}, {} bytes)",
                    c.index, c.name, c.arity, c.code_len
                );
            }
        }
    }
}

fn reason_str(r: &PauseReason) -> String {
    match r {
        PauseReason::Entry => "entry".into(),
        PauseReason::Breakpoint { id } => format!("breakpoint #{id}"),
        PauseReason::Step => "step".into(),
        PauseReason::Interrupt => "interrupt".into(),
        PauseReason::Watchpoint { id } => format!("watchpoint #{id}"),
        PauseReason::Exception { uncaught } => {
            if *uncaught {
                "uncaught exception".into()
            } else {
                "exception".into()
            }
        }
        PauseReason::HostCall { name } => format!("host call {name}"),
    }
}

fn fmt_location(l: &Location) -> String {
    let line = l.line.map(|n| format!(" line {n}")).unwrap_or_default();
    format!("{}@{:04}{}", l.chunk_name, l.ip, line)
}

fn fmt_frame(f: &FrameInfo) -> String {
    let line = f.line.map(|n| format!(" line {n}")).unwrap_or_default();
    format!("#{} {}@{:04}{}", f.depth, f.chunk_name, f.ip, line)
}

fn print_help() {
    let memory_limit = vybe_runtime::debugger::MAX_DEBUG_MEMORY_READ;
    eprintln!(
        "  control:  c continue · s step-in · n step-over · o step-out · si stepi · q quit\n\
         \x20 breaks:   b <line> · b <file>:<line> · b <chunk>@<offset> · [if <cond>] · bl · bd <id> · enable/disable\n\
         \x20 breaks+:  bf <fn> · bh <module>:<name> [invalid]|off · lp <line> <msg with {{expr}}> logpoint · rt <line> run-to · ignore <id> <n> · catch throw|uncaught|off\n\
         \x20 data:     wp <name> watchpoint · wps list · unwp clear · fibers/threads · restart\n\
         \x20 inspect:  bt backtrace · locals [frame] · stack · g/globals [prefix] · dis [@offset] [n] · targets <offset> · chunks\n\
         \x20 loader:   trace-global-inits on|off · bgi <global-name|*|off>  break before init\n\
         \x20 frame:    fr/frame [n]  EVERY slot incl. compiler+capture, with local_count/capture_base\n\
         \x20 profile:  report [filter|reset-samples]  timings and instruction/host samples\n\
         \x20 vars:     p <name>[.field][idx] or p <expr> · set <name> = <literal> · watch <expr> · watches · unwatch\n\
         \x20 memory:   memory <index> <address> [byte-count]  live read-only hex/ASCII dump (default 64, max {memory_limit})\n\
         \x20 gui:      widgets/controls · click <control> · fire <control> <event> · close [control]\n\
         \x20 input:    key <keydown|keyup> <key> <keyCode> [code]  browser keyboard dispatch\n\
         \x20 pointer:  pointer <mousedown|mouseup|mousemove> <x> <y> [button]  browser hit-testing\n\
         \x20 gui+:     draws [control]    canvas ink + bounds · capture [control] [file.png] offscreen PNG\n\
         \x20 stream+:  trace canvas on|off  (draw routing — which control each draw resolved to)\n\
         \x20 reload:   reload  (recompile + swap changed fn bodies in place; heap/globals kept)\n\
         \x20 stream:   trace on|off  (live opcode stream — the VYBE_TRACE replacement)\n\
         \x20 chunk may be a name or a numeric index (see `chunks`)."
    );
}

#[cfg(test)]
mod tests {
    use super::{memory_dump_rows, parse_command, parse_memory_arguments, reset_samples};
    use std::sync::{Arc, Mutex, atomic::Ordering};
    use std::time::{Duration, Instant};
    use vybe_runtime::debugger::DebugReport;

    #[test]
    fn memory_arguments_preserve_exact_addresses_and_indices() {
        assert_eq!(parse_memory_arguments(&["0", "34731088"]).unwrap(), (0, 34731088, 64));
        assert_eq!(parse_memory_arguments(&["2", "0x100000001", "0x20"]).unwrap(), (2, 0x100000001, 32));
        assert_eq!(parse_memory_arguments(&["0", "18446744073709551615", "0"]).unwrap(), (0, u64::MAX, 0));
        for parts in [vec![], vec!["0"], vec!["0", "1", "2", "3"],
            vec!["-1", "0"], vec!["4294967296", "0"], vec!["0", "1.5"],
            vec!["0", "18446744073709551616"], vec!["0", "1", "4294967296"]] {
            assert!(parse_memory_arguments(&parts).is_err(), "{parts:?}");
        }
    }

    #[test]
    fn memory_command_and_dump_are_bounded_and_lossless() {
        assert!(matches!(parse_command("mem 2 0x100000001 32").unwrap(),
            vybe_runtime::DebugCommand::ReadMemory { memory: 2, address: 0x100000001, length: 32 }));
        let bytes: Vec<u8> = (0..17).collect();
        let rows = memory_dump_rows(0x100000001, &bytes);
        assert_eq!(rows.len(), 2);
        assert!(rows[0].starts_with("0000000100000001  00 01 02 03"));
        assert!(rows[1].starts_with("0000000100000011  10"));
        assert!(memory_dump_rows(0, b"A \0\xff")[0].ends_with("|A ..|"));
        assert!(memory_dump_rows(0, &[]).is_empty());
    }

    #[test]
    fn sample_reset_preserves_execution_and_timing_state() {
        let mut state = DebugReport::default();
        state.function_samples.insert("draw".into(), 12);
        state.line_samples.insert("draw:42".into(), 8);
        state.host_samples.insert(Arc::from("putImageData"), (2, Duration::from_millis(4)));
        state.instructions = 4096;
        state.live.instructions.store(4100, Ordering::Relaxed);
        state.location = Some("draw@16".into());
        state.host_sample_tick = 17;
        state.paused = Duration::from_secs(2);
        state.completed.push(("compile".into(), Duration::from_secs(3)));
        state.active_host = Some((Arc::from("read"), Instant::now()));
        let report = Arc::new(Mutex::new(state));

        assert_eq!(reset_samples(&report), 4100);
        let state = report.lock().unwrap();
        assert!(state.function_samples.is_empty());
        assert!(state.line_samples.is_empty());
        assert!(state.host_samples.is_empty());
        assert_eq!(state.instructions, 4096);
        assert_eq!(state.live.instructions.load(Ordering::Relaxed), 4100);
        assert_eq!(state.location.as_deref(), Some("draw@16"));
        assert_eq!(state.host_sample_tick, 17);
        assert_eq!(state.paused, Duration::from_secs(2));
        assert_eq!(state.completed, vec![("compile".into(), Duration::from_secs(3))]);
        assert_eq!(state.active_host.as_ref().unwrap().0.as_ref(), "read");
    }
}
