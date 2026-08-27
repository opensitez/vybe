//! The document the window shows — the one tree every frontend builds.
//!
//! `compiler::primitives::gui` is the single GUI emit layer for every language,
//! and it lowers control creation to `web:dom.createElement`, control
//! properties to `web:dom` / `web:html` / `web:cssom`, and `OnClick := h` to
//! `addEventListener("click", h)`. VCL, WinForms, Flutter and SDL all end up in
//! the SAME `widgets` document, so nothing in this module is
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

use vybe_platform_web::engine::{DocumentId, NodeId};
use vybe_platform_web::html;
use vybe_runtime::Value;
use widgets::LayoutRect;
use widgets::dom::Document;

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
        html::active_document(),
        std::sync::atomic::Ordering::Relaxed,
    );
}

static PINNED: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

/// Borrow the document, but only if the guest actually built its UI in it.
///
/// `html::active_document()` CREATES a document on first call, so its existence
/// proves nothing: a designer form that never touched `web:*` would get a fresh
/// empty one and we would happily route the mouse at it. Having controls is the
/// test, and it is the same question `render_into` asks before it picks a form
/// to paint.
pub fn with_live<T>(f: impl FnOnce(&mut Document) -> T) -> Option<T> {
    widgets::dom::with_document(active(), |document| {
        (document.form().control_count() > 0).then(|| f(document))
    })
    .flatten()
}

/// Did the guest build a UI in its document?
///
/// The same test [`with_live`] gates on, asked without borrowing: a document
/// with content is a running one. This is what tells the runner to present a
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
    use super::{NodeId, with_live};
    use widgets::dom::Document;

    /// The INDENTED form, which is what a person reading a dump wants.
    ///
    /// `outer_html` is the DOM getter and adds no whitespace, as the HTML
    /// fragment serialization algorithm requires — correct as markup, one long
    /// line as a debugger's output. The two callers want different things and
    /// now ask for them by name.
    pub fn outer_html(node: NodeId) -> Option<String> {
        with_live(|d| d.outer_html_pretty(node))
    }

    pub fn style(node: NodeId, property: &str) -> Option<String> {
        with_live(|d| d.get_style_property(node, property))
    }

    pub fn set_style(node: NodeId, property: &str, value: &str) -> Option<()> {
        with_live(|d| d.set_style_property(node, property, value))
    }

    /// Every declaration on the element, in the order they serialise.
    pub fn declarations(node: NodeId) -> Option<Vec<(String, String)>> {
        with_live(|d: &mut Document| {
            d.style(node)
                .map(|s| {
                    s.iter()
                        .map(|(k, v)| (k.to_string(), v.to_string()))
                        .collect()
                })
                .unwrap_or_default()
        })
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
        with_live(|d| d.text_content(node))
    }

    pub fn set_text(node: NodeId, value: &str) -> Option<()> {
        with_live(|d| d.set_text_content(node, value))
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
/// `with_live` borrows `widgets::dom` directly, which is the toolkit
/// whether or not the toolkit is the live engine. Every GUI command in the step
/// debugger went through it, so under `--engine webcore` they reported an empty
/// tree for a document webcore had built perfectly well — the debugger was
/// inspecting the engine that was NOT running.
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
    // `window.innerWidth` / `innerHeight` — asked through the seam, in the
    // vocabulary both engines already answer, rather than by borrowing the
    // toolkit's document. Under `--engine webcore` the toolkit's is empty, so
    // this reported "no live document to capture" for a form webcore had laid
    // out perfectly well.
    if !vybe_platform_web::present::has_content(active()) {
        return None;
    }
    match vybe_platform_web::engine::window(vybe_platform_web::engine::WindowOp::InnerSize(
        active(),
    )) {
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
    pub rect: Option<LayoutRect>,
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

/// Every element the guest put in the document, with its geometry, its
/// observable properties and its wired listeners.
/// ⛔ Through the SEAM, not through `widgets`' document.
///
/// This walked the toolkit's tree directly, which is the toolkit whether or
/// not the toolkit is the live engine — so under `--engine webcore` it saw an
/// empty form and every caller inherited that: `--capture-control` answered
/// "no control named X" for controls webcore had laid out perfectly well, and
/// the debugger's control list came back blank. `node_by_id` and `has_content`
/// were moved for the same reason; this was the last one reaching past the
/// surface that owns the document.
pub fn controls() -> Vec<DomControl> {
    use vybe_platform_web::engine::{DomOp, DomValue, apply};

    let document = active();
    let mut listener_types: std::collections::HashMap<NodeId, Vec<String>> =
        std::collections::HashMap::new();
    // Listeners stay with the toolkit on purpose: `EventOp` is host
    // bookkeeping that BOTH engines delegate to it, so this is the one place
    // where the toolkit IS the authority.
    for (node, kind, _) in html::document_listeners(document) {
        listener_types.entry(node).or_default().push(kind);
    }

    let text_of = |node: NodeId, op: DomOp| match apply(document, op) {
        DomValue::Text(s) => s,
        _ => {
            let _ = node;
            String::new()
        }
    };

    // Every element, not only the named ones — `Name` is optional, and a form
    // that builds its buttons in a loop has none.
    //
    // ⛔ `AllElements`, NOT `querySelectorAll("*")`. A selector matches only
    // what is IN the document, so a created-and-never-appended control would
    // vanish from this list — and `DomControl::connected` exists precisely to
    // report that control. Listing only the tree turns "you forgot
    // `Controls.Add`" into "no control named X", which is a different bug.
    let DomValue::Nodes(nodes) = apply(document, DomOp::AllElements) else {
        return Vec::new();
    };
    let mut out = Vec::new();
    for node in nodes {
        let tag = text_of(node, DomOp::NodeName(node)).to_ascii_lowercase();
        // The document's own skeleton is not a control, and listing it would
        // bury a form's five buttons under the page that contains them.
        //
        // ⛔ `body` is NOT in this list. A form IS the body — `emit_control_element`
        // special-cases the tag for exactly that reason — so a designer's
        // `Form1` is the body element and filtering it out would make the whole
        // form unaddressable by its own name.
        if matches!(tag.as_str(), "html" | "head" | "title" | "style" | "meta" | "script" | "link") {
            continue;
        }
        let id = text_of(node, DomOp::GetAttribute(node, "id".into()));
        // A zero-area box is "present and never laid out", which is a
        // different bug from "never appended" — `connected` separates them,
        // and `control_rect` says which one it found.
        let rect = match apply(document, DomOp::BoundingClientRect(node)) {
            DomValue::Rect { x, y, width, height } if width > 0.0 || height > 0.0 => {
                Some(LayoutRect::new(x as f32, y as f32, width as f32, height as f32))
            }
            _ => None,
        };
        let mut properties = Vec::new();
        let text = text_of(node, DomOp::TextContent(node));
        if !text.is_empty() {
            properties.push(("textContent".to_string(), text));
        }
        let value = text_of(node, DomOp::Value(node));
        if !value.is_empty() {
            properties.push(("value".to_string(), value));
        }
        if matches!(apply(document, DomOp::Checked(node)), DomValue::Bool(true)) {
            properties.push(("checked".to_string(), "true".to_string()));
        }
        for css in ["left", "top", "width", "height"] {
            let v = text_of(node, DomOp::GetStyleProperty(node, css.to_string()));
            if !v.is_empty() {
                properties.push((css.to_string(), v));
            }
        }
        let mut events = listener_types.get(&node).cloned().unwrap_or_default();
        events.sort();
        let connected = matches!(apply(document, DomOp::IsConnected(node)), DomValue::Bool(true));
        out.push(DomControl {
            node,
            id,
            tag,
            rect,
            connected,
            properties,
            events,
        });
    }
    out
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
    html::event_object(&kind.to_ascii_lowercase(), target)
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
        // operation `platforms/web` already exposes, and reaching past it into
        // `widgets::dom` made this crate a second driver of a document
        // web is supposed to own.
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
