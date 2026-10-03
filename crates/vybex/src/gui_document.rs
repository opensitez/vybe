//! The document the window shows — the one tree every frontend builds.
//!
//! `compiler::primitives::gui` is the single GUI emit layer for every language,
//! and it lowers control creation to `web:dom.createElement`, control
//! properties to `web:dom` / `web:html` / `web:cssom`, and `OnClick := h` to
//! `addEventListener("click", h)`. VCL, WinForms, Flutter and SDL all end up in
//! the same active browser document, so nothing in this module is
//! framework-specific and nothing in it may become so.
//!
//! `GuiState` still owns what is not a DOM element — form lifecycle flags,
//! dialogs, timers, overlay canvases — and a designer form built straight into
//! `GuiState.form` (`launch_vybewidget_form`) never opens a document at all.
//! The document is therefore used WHEN IT HAS CONTENT, with `GuiState` as the
//! fallback: the same rule `gui_capture::render_into` already paints by.
//!
//! ## Why a module rather than code in `gui_launch`
//!
//! Three callers need the same answers: the window runner (route input, drain
//! events), the step debugger's `widgets` dump, and the debugger's
//! `click`/`fire` hook. Deriving "which tree is live" separately in each is how
//! they drift apart — the window rendering the document while the debugger
//! reports on an empty `GuiState` is exactly the state this fixes.

use vybe_platform_web::engine::{DocumentId, NodeId, UiEventFields};
use vybe_platform_web::html;
use vybe_runtime::Value;

pub(crate) fn fn_arity(val: &Value) -> usize {
    match val {
        Value::Object(object) => match &object.lock().unwrap().kind {
            vybe_runtime::value::ObjectKind::Function(function) => function.arity as usize,
            _ => 0,
        },
        _ => 0,
    }
}

/// This agent's ambient document — `window.document`.
///
/// `html::active_document()` is thread-local, and correctly so: a document
/// belongs to an AGENT, and two guests running side by side must not share one.
/// The debugger, though, reads the guest's document from its own REPL thread —
/// where that thread-local is a different, empty document. [`pin`] records the
/// guest's handle so a reader outside the guest's thread asks about the tree
/// the guest actually built.
pub fn active() -> DocumentId {
    match PINNED.load(std::sync::atomic::Ordering::Relaxed) {
        0 => html::active_document(),
        id => id,
    }
}

/// Pin the calling thread's document as the one every reader means.
///
/// Called on the VM's thread while attaching the debugger — before the guest
/// runs, so the handle it opens here IS the one the guest goes on to use.
pub fn pin() {
    PINNED.store(
        html::active_document_for_debugger(),
        std::sync::atomic::Ordering::Relaxed,
    );
}

static PINNED: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

/// Did the guest build a UI in its document?
///
/// A document with content is a running one. This tells the runner to present a
/// window for a program that never asked to be run — which is every frontend,
/// since a page is not told to run.
///
/// A program that declares [`vybe_ast::AppShell::Windowed`] is presented even
/// when this is false: the declaration covers a UI built later, from a timer or
/// a handler, which no test at this instant can see.
pub fn has_content() -> bool {
    // Through `platforms/web`, so the answer is about the LIVE engine's
    // document. `with_live` below still names the toolkit, which is why this
    // no longer goes through it: under `--engine webcore` that asks the wrong
    // tree and reports an empty UI for a form webcore laid out perfectly well.
    vybe_platform_web::present::has_content(active())
}

/// Read and write one element, live — the inspector half of the debugger.
///
/// Every one of these goes through the SAME `Document` entry point the guest
/// uses (`set_style_property`, `set_attribute`, `set_text_content`), so what
/// the inspector does to an element is exactly what a program doing it would
/// do. An inspector with its own write path would be able to produce states a
/// program cannot reach, and would then "prove" behaviour that never happens in
/// a real run — the same mistake as a probe that calls a handler directly.
pub mod inspect {
    use super::NodeId;
    use vybe_platform_web::engine::{self, DomOp, DomValue};

    /// The INDENTED form, which is what a person reading a dump wants.
    ///
    /// `outer_html` is the DOM getter and adds no whitespace, as the HTML
    /// fragment serialization algorithm requires — correct as markup, one long
    /// line as a debugger's output. The two callers want different things and
    /// now ask for them by name.
    pub fn outer_html(node: NodeId) -> Option<String> {
        match engine::apply(super::active(), DomOp::OuterHtml(node)) {
            DomValue::Text(html) => Some(html),
            _ => None,
        }
    }

    pub fn style(node: NodeId, property: &str) -> Option<String> {
        match engine::apply(super::active(), DomOp::GetStyleProperty(node, property.into())) {
            DomValue::Text(value) => Some(value),
            _ => None,
        }
    }

    pub fn set_style(node: NodeId, property: &str, value: &str) -> Option<()> {
        engine::apply(super::active(), DomOp::SetStyleProperty(node, property.into(), value.into()));
        Some(())
    }

    /// Every declaration on the element, in the order they serialise.
    pub fn declarations(node: NodeId) -> Option<Vec<(String, String)>> {
        match engine::apply(super::active(), DomOp::StyleDeclarations(node)) {
            DomValue::Properties(properties) => Some(properties),
            _ => None,
        }
    }

    pub fn attribute(node: NodeId, name: &str) -> Option<String> {
        // Seam, not toolkit — see `node_by_id`.
        match vybe_platform_web::engine::apply(
            super::active(),
            vybe_platform_web::engine::DomOp::GetAttribute(node, name.to_string()),
        ) {
            vybe_platform_web::engine::DomValue::Text(v) => Some(v),
            _ => None,
        }
    }

    pub fn set_attribute(node: NodeId, name: &str, value: &str) -> Option<()> {
        vybe_platform_web::engine::apply(
            super::active(),
            vybe_platform_web::engine::DomOp::SetAttribute(
                node,
                name.to_string(),
                value.to_string(),
            ),
        );
        Some(())
    }

    pub fn text(node: NodeId) -> Option<String> {
        match engine::apply(super::active(), DomOp::TextContent(node)) {
            DomValue::Text(value) => Some(value),
            _ => None,
        }
    }

    pub fn set_text(node: NodeId, value: &str) -> Option<()> {
        engine::apply(super::active(), DomOp::SetTextContent(node, value.into()));
        Some(())
    }
}

/// The live tree serialised as HTML, when the document is the live tree.
///
/// `controls()` reports each element's *properties*; this reports the
/// **structure** — what is inside what, with which tag. They answer different
/// questions and a rendering bug is usually one or the other: a control with
/// the right properties in the wrong parent looks fine in a property dump.
///
/// It is also the only form of this evidence that diffs. A PNG hash says
/// "something moved"; this says which element, and a golden file can be
/// reviewed in a patch.
pub fn html() -> Option<String> {
    let markup = engine_html();
    (!markup.is_empty()).then_some(markup)
}

/// The live tree as the ENGINE has it — through the seam, so it answers for
/// whichever engine is installed.
///
/// The step debugger asks the selected engine for this tree, so it sees the
/// same document the user sees.
///
/// `outerHTML` on the document is the one question that needs no toolkit type
/// to answer, which is why the structure dump is the first of them to move.
pub fn engine_html() -> String {
    match vybe_platform_web::engine::apply(
        active(),
        vybe_platform_web::engine::DomOp::OuterHtml(vybe_platform_web::engine::DOCUMENT),
    ) {
        vybe_platform_web::engine::DomValue::Text(markup) => markup,
        _ => String::new(),
    }
}

/// The size the document says the window is, when the document is the live
/// tree.
///
/// A form's `Width`/`Height` are CSS on the body — `primitives/gui.rs` lowers
/// them to `web:cssom.setStyleProperty` like any other geometry — so the
/// document's viewport, not `GuiState`, is what a program that sets them
/// actually set. `GuiState.width`/`height` keep their defaults and a 280×400
/// form opened at 800×600.
pub fn viewport() -> Option<(u32, u32)> {
    // `window.innerWidth` / `innerHeight` come from the active engine.
    if !vybe_platform_web::present::has_content(active()) {
        return None;
    }
    match vybe_platform_web::engine::window(
        vybe_platform_web::engine::WindowOp::InnerSize(active()),
    ) {
        vybe_platform_web::engine::WindowValue::Pair(w, h) if w >= 1.0 && h >= 1.0 => {
            Some((w.round() as u32, h.round() as u32))
        }
        _ => None,
    }
}

/// One realized element, in the vocabulary a GUI debugger reports in.
pub struct DomControl {
    pub node: NodeId,
    /// The `id` attribute — what every frontend lowers a control's `Name` to,
    /// and therefore the name a user types at the debugger.
    pub id: String,
    pub tag: String,
    /// The LAID-OUT rect, looked up in the FORM's tree.
    ///
    /// `None` means the element is not in that tree — which is usually not
    /// "nothing has laid out yet" but **"this element was never appended"**.
    /// A created-and-never-inserted control is styled, named, addressable and
    /// absent, and it is the single most common way a GUI silently renders
    /// nothing. [`DomControl::connected`] separates the two.
    pub rect: Option<DomRect>,
    /// Is the element actually in the document?
    ///
    /// Distinguishes "created and never appended" from "appended but not yet
    /// laid out". Both show no rect and they are completely different bugs:
    /// the first is a missing `appendChild`, the second is a missing layout
    /// pass.
    pub connected: bool,
    pub properties: Vec<(String, String)>,
    /// Registered listener types, in DOM spelling (`click`, `input`, …).
    pub events: Vec<String>,
}

#[derive(Clone, Copy)]
pub struct DomRect {
    pub x: f32,
    pub y: f32,
    pub w: f32,
    pub h: f32,
}

/// Every element the guest put in the document, with its geometry, its
/// observable properties and its wired listeners.
pub fn controls() -> Vec<DomControl> {
    use vybe_platform_web::engine::{DomOp, DomValue, apply};
    let document = active();
    let mut listener_types: std::collections::HashMap<NodeId, Vec<String>> =
        std::collections::HashMap::new();
    for (node, kind, _) in html::document_listeners(document) {
        listener_types.entry(node).or_default().push(kind);
    }
    let DomValue::Nodes(nodes) = apply(document, DomOp::QuerySelectorAll("*".into())) else {
        return Vec::new();
    };
    nodes.into_iter().map(|node| {
            let text = |op| match apply(document, op) {
                DomValue::Text(value) => value,
                _ => String::new(),
            };
            let id = text(DomOp::GetAttribute(node, "id".into()));
            let tag = text(DomOp::NodeName(node));
            let connected = matches!(apply(document, DomOp::IsConnected(node)), DomValue::Bool(true));
            let rect = match apply(document, DomOp::BoundingClientRect(node)) {
                DomValue::Rect { x, y, width, height } if connected => Some(DomRect {
                    x: x as f32, y: y as f32, w: width as f32, h: height as f32,
                }),
                _ => None,
            };
            let mut properties = Vec::new();
            let content = text(DomOp::TextContent(node));
            if !content.is_empty() {
                properties.push(("textContent".to_string(), content));
            }
            let value = text(DomOp::Value(node));
            if !value.is_empty() {
                properties.push(("value".to_string(), value));
            }
            if matches!(apply(document, DomOp::Checked(node)), DomValue::Bool(true)) {
                properties.push(("checked".to_string(), "true".to_string()));
            }
            for css in ["left", "top", "width", "height"] {
                let v = text(DomOp::GetStyleProperty(node, css.into()));
                if !v.is_empty() {
                    properties.push((css.to_string(), v));
                }
            }
            let mut events = listener_types.get(&node).cloned().unwrap_or_default();
            events.sort();
            DomControl {
                node,
                id,
                tag,
                rect,
                connected,
                properties,
                events,
            }
    }).collect()
}

/// Resolve a control name the way a debugger user types it — the `id`
/// attribute, case-insensitively, because that is how every `GuiState` caller
/// has always addressed a control.
///
/// `n<node>` also resolves, so a control the author never named is still
/// addressable: that is the handle `widgets` prints for it, and without it a
/// form that builds its buttons in a loop can be listed but never clicked.
pub fn node_by_id(name: &str) -> Option<NodeId> {
    // Through the seam, so the inspector answers for whichever engine is live.
    // This walked `widgets`' document directly, which is the toolkit
    // whether or not the toolkit is running — so under `--engine webcore`
    // every `css`/`attr`/`text` command reported "no control named X" for
    // controls that were plainly in the tree. `html` was fixed first and made
    // that obvious: the dump listed the element the next command denied.
    use vybe_platform_web::engine::{DomOp, DomValue, apply};
    match apply(active(), DomOp::GetElementById(name.to_string())) {
        DomValue::Node(node) => return Some(node),
        _ => {}
    }
    // `getElementById` is case-SENSITIVE (DOM §4.5), and a debugger user types
    // what they saw. Fall back to a case-insensitive sweep of the ids that
    // exist, which is a convenience of this command and not of the DOM.
    if let DomValue::Nodes(all) = apply(active(), DomOp::QuerySelectorAll("[id]".into())) {
        for node in all {
            if let DomValue::Text(id) = apply(active(), DomOp::GetAttribute(node, "id".into())) {
                if id.eq_ignore_ascii_case(name) {
                    return Some(node);
                }
            }
        }
    }
    // `n<id>` — the internal node number, for anything the author never named.
    name.strip_prefix('n')?.parse().ok()
}

/// The listeners registered for one event type on one node, in registration
/// order. `kind` is accepted in any case: `Click` from a debugger command and
/// `click` from the DOM are the same event.
pub fn listeners_for(node: NodeId, kind: &str) -> Vec<Value> {
    html::listeners_for(active(), node, kind)
}

/// The `Event` a synthesised dispatch hands its listener — the same object the
/// drained path builds, so a simulated click is indistinguishable from a real
/// one to the handler.
pub fn event_object(kind: &str, target: NodeId) -> Value {
    html::event_object(
        &UiEventFields { kind: kind.to_ascii_lowercase(), ..Default::default() },
        target,
    )
}

/// One drained interaction, ready to hand to the VM.
pub struct Dispatch {
    pub callback: Value,
    /// The `Event` object the listener receives — its only argument.
    pub event: Value,
    /// DOM event type — `click`, `input`, `change`, …
    pub kind: String,
    /// The `id` of the element the event targeted; empty for the body/form.
    /// Reported for tracing; the listener reads it off `event.target`.
    pub sender: String,
}

/// Drain what the user did into calls waiting to be made.
///
/// The document lock is released before this returns, and deliberately so: a
/// handler runs `web:dom` host calls of its own, and those re-enter the very
/// mutex `with_live` holds. Everything a dispatch needs is resolved here, up
/// front, so the caller invokes with no lock held.
pub fn drain() -> Vec<Dispatch> {
    let document = active();
    let pending = html::pending_dispatches(document);
    if pending.is_empty() {
        return Vec::new();
    }
    let mut out = Vec::new();
    for (callback, event) in pending {
        let kind = event_field(&event, "type")
            .map(|v| v.as_str().to_string())
            .unwrap_or_default();
        let node = event_field(&event, "target")
            .map(|v| v.as_f64() as NodeId)
            .unwrap_or(0);
        // Through the web surface, not around it. `getAttribute` is a DOM
        // operation `platforms/web` already exposes.
        let sender = match vybe_platform_web::engine::apply(
            document,
            vybe_platform_web::engine::DomOp::GetAttribute(node, "id".to_string()),
        ) {
            vybe_platform_web::engine::DomValue::Text(id) => id,
            _ => String::new(),
        };
        out.push(Dispatch {
            callback,
            event,
            kind,
            sender,
        });
    }
    out
}

fn event_field(event: &Value, key: &str) -> Option<Value> {
    match event {
        Value::Object(obj) => obj.lock().ok()?.properties.get(key).cloned(),
        _ => None,
    }
}
