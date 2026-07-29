//! Built-in step-debugger REPL — the client half of the debug surface. The VM
//! (in `vybe_runtime`) owns the pause/breakpoint state machine; this module is
//! pure transport + presentation: it attaches channels to the VM, spawns a
//! stdin-reader thread and an event-printer thread, and formats the typed
//! protocol for a terminal.
//!
//! The VM stays on the main thread (it is not `Send`); these worker threads hold
//! only channel endpoints, exactly like the browser debug server pattern.

use std::io::{BufRead, Write};
use std::sync::mpsc::{Sender, channel};
use std::sync::{Arc, Mutex};
use std::thread;

use vybe_runtime::debugger::{
    ChunkRef, DebugEvent, DebugResponse, FrameInfo, Location, PauseReason,
};
use vybe_runtime::{DebugCommand, DebugRequest, VM};
use vybe_platform_vybe::gui_state::GuiState;

/// Attach a debugger to `vm` and spawn the REPL worker threads. Call this before
/// running the VM; it pauses on entry so breakpoints can be set first. `gui` is
/// the live GUI state shared with the host functions — the `widgets` command
/// reads it directly (client-side), so it reflects live controls regardless of
/// the isolated eval VM.
pub fn attach(vm: &mut VM, gui: Arc<Mutex<GuiState>>) {
    let (cmd_tx, cmd_rx) = channel::<DebugRequest>();
    let (evt_tx, evt_rx) = channel::<DebugEvent>();
    vm.attach_debugger(cmd_rx, evt_tx, /* pause_on_entry */ true);

    // Event printer: renders async events (paused / resumed / exited / opcode).
    thread::spawn(move || {
        for event in evt_rx {
            print_event(&event);
        }
    });

    // Stdin reader: parse a line → command → send → print the reply.
    thread::spawn(move || {
        let stdin = std::io::stdin();
        let mut lines = stdin.lock().lines();
        banner();
        loop {
            prompt();
            let Some(Ok(line)) = lines.next() else { break };
            let line = line.trim();
            if line.is_empty() {
                continue;
            }
            // `widgets`/`gui`/`controls` are handled entirely client-side: they
            // read the live GuiState Arc directly (safe to lock while the VM is
            // paused or running), so they never round-trip through the VM.
            let head = line.split_whitespace().next().unwrap_or("");
            if matches!(head, "widgets" | "controls") {
                print_widgets(&gui);
                continue;
            }
            match parse_command(line) {
                Ok(command) => {
                    if !send_and_print(&cmd_tx, command) {
                        break; // channel closed — VM gone
                    }
                }
                Err(msg) => eprintln!("  {msg}"),
            }
        }
    });
}

/// Dump the live GUI state (controls, their properties, wired events). Reads the
/// shared `GuiState` directly — reflects the live program, not the eval VM.
fn print_widgets(gui: &Arc<Mutex<GuiState>>) {
    let g = match gui.lock() {
        Ok(g) => g,
        Err(_) => {
            eprintln!("  (gui state unavailable)");
            return;
        }
    };
    if g.control_names.is_empty() {
        eprintln!("  (no controls realized yet)");
        return;
    }
    eprintln!(
        "  form {}×{}  running={}  ({} control(s))",
        g.width,
        g.height,
        g.should_run,
        g.control_names.len()
    );
    for name in &g.control_names {
        // Properties recorded for this control (keyed by (control, prop_lower)).
        let mut props: Vec<String> = g
            .properties
            .iter()
            .filter(|((c, _), _)| c.eq_ignore_ascii_case(name))
            .map(|((_, p), v)| format!("{p}={v}"))
            .collect();
        props.sort();
        // Events wired on this control (keys are "control.event").
        let mut events: Vec<String> = g
            .event_handlers
            .keys()
            .filter_map(|k| k.rsplit_once('.'))
            .filter(|(c, _)| c.eq_ignore_ascii_case(name))
            .map(|(_, ev)| ev.to_string())
            .collect();
        events.sort();
        let prop_str = if props.is_empty() { String::new() } else { format!("  {{{}}}", props.join(", ")) };
        let evt_str = if events.is_empty() { String::new() } else { format!("  events[{}]", events.join(",")) };
        eprintln!("  • {name}{prop_str}{evt_str}");
    }
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

fn parse_command(line: &str) -> Result<DebugCommand, String> {
    let mut parts = line.split_whitespace();
    let cmd = parts.next().unwrap_or("");
    let rest: Vec<&str> = parts.collect();
    Ok(match cmd {
        "c" | "cont" | "continue" => DebugCommand::Continue,
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
        "dis" | "disasm" => DebugCommand::Disasm {
            window: rest.first().and_then(|s| s.parse().ok()).unwrap_or(4),
        },
        "chunks" => DebugCommand::Chunks,
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
            control: rest.first().ok_or("usage: click <control>  (see `widgets` for names)")?.to_string(),
            event: "Click".to_string(),
        },
        "fire" => {
            let control = rest.first().ok_or("usage: fire <control> <event>")?.to_string();
            let event = rest.get(1).ok_or("usage: fire <control> <event>")?.to_string();
            DebugCommand::FireEvent { control, event }
        }
        "close" | "window-close" => DebugCommand::FireEvent {
            control: rest.first().unwrap_or(&"form").to_string(),
            event: "Close".to_string(),
        },

        // ── function breakpoint / logpoint / run-to-cursor / ignore ──
        "bf" | "break-fn" => {
            let name = rest.first().ok_or("usage: bf <function> [if <cond>]")?.to_string();
            let condition = rest
                .iter()
                .position(|t| *t == "if")
                .map(|i| rest[i + 1..].join(" "))
                .filter(|c| !c.trim().is_empty());
            DebugCommand::BreakFunction { name, condition }
        }
        "logpoint" | "lp" => {
            let line = rest.first().ok_or("usage: lp <line> <message>")?.parse().map_err(|_| "bad line")?;
            let message = rest[1..].join(" ");
            if message.is_empty() {
                return Err("usage: lp <line> <message>  (use {expr} to interpolate)".into());
            }
            DebugCommand::Logpoint { line, message }
        }
        "runto" | "rt" | "tbreak" => {
            let line = rest.first().ok_or("usage: rt <line>")?.parse().map_err(|_| "bad line")?;
            DebugCommand::RunToLine { line }
        }
        "ignore" => {
            let id = rest.first().ok_or("usage: ignore <id> <count>")?.parse().map_err(|_| "bad id")?;
            let count = rest.get(1).ok_or("usage: ignore <id> <count>")?.parse().map_err(|_| "bad count")?;
            DebugCommand::SetIgnoreCount { id, count }
        }
        // ── exception breakpoints ──
        "catch" => match rest.first().copied() {
            Some("throw") => DebugCommand::ExceptionBreak { on_throw: true, on_uncaught: false },
            Some("uncaught") => DebugCommand::ExceptionBreak { on_throw: false, on_uncaught: true },
            Some("off") | None => DebugCommand::ExceptionBreak { on_throw: false, on_uncaught: false },
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
    let spec = rest
        .first()
        .ok_or("usage: b <chunk>:<line> [if <cond>]")?;
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

fn print_response(resp: &DebugResponse) {
    match resp {
        DebugResponse::Ok => {}
        DebugResponse::Error(e) => eprintln!("  error: {e}"),
        DebugResponse::Value(s) => eprintln!("  {s}"),
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
            if *uncaught { "uncaught exception".into() } else { "exception".into() }
        }
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
    eprintln!(
        "  control:  c continue · s step-in · n step-over · o step-out · si stepi · q quit\n\
         \x20 breaks:   b <line> · b <file>:<line> · b <chunk>@<offset> · [if <cond>] · bl · bd <id> · enable/disable\n\
         \x20 breaks+:  bf <fn> · lp <line> <msg with {{expr}}> logpoint · rt <line> run-to · ignore <id> <n> · catch throw|uncaught|off\n\
         \x20 data:     wp <name> watchpoint · wps list · unwp clear · fibers/threads · restart\n\
         \x20 inspect:  bt backtrace · locals [frame] · stack · g/globals [prefix] · dis [n] · chunks\n\
         \x20 vars:     p <name>[.field][idx] or p <expr> · set <name> = <literal> · watch <expr> · watches · unwatch\n\
         \x20 gui:      widgets/controls · click <control> · fire <control> <event> · close [control]\n\
         \x20 reload:   reload  (recompile + swap changed fn bodies in place; heap/globals kept)\n\
         \x20 stream:   trace on|off  (live opcode stream — the VYBE_TRACE replacement)\n\
         \x20 chunk may be a name or a numeric index (see `chunks`)."
    );
}
