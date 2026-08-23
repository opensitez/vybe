//! htmlbox as the engine behind `web:*`.
//!
//! The sibling of `engine_widgets.rs`, against the same trait. `engine.rs`
//! names neither engine, so which one is live is decided by which `install()`
//! runs — that is what makes them swappable.
//!
//! WHAT THIS FILE OWNS AND WHAT IT DELEGATES
//!
//! `DocumentId` is ONE namespace shared by `document()` and `window()`:
//! `WindowOp::Open`, `Document`, `DefaultView` and `AdoptTopLevel` all mint or
//! return one. If `vybe_widgets` answered those while htmlbox answered
//! `document()`, the ids handed back would point into the widget document
//! table and every following `apply()` would miss — two document tables over
//! one id space. So htmlbox owns the whole browsing context: all of `DomOp`
//! and the id-minting `WindowOp`s.
//!
//! `EventOp` is NOT among them, which looks wrong until you follow the two
//! queues. `DomOp::DrainEvents` is per-document and carries DOM listener
//! events. `EventOp::Poll`/`Pending` read `vybe_widgets::ui_events::queue()`,
//! a PROCESS-WIDE queue of raw input that the winit shell in `gui_launch.rs`
//! pushes into. Different queues, different producers — so `EventOp`
//! delegates, and swapping the DOM engine does not disturb input delivery.
//!
//! What else is delegated is the part that is not DOM at all: the scheduler,
//! window geometry and chrome, and pointer state. Host bookkeeping, identical
//! under either engine.
//!
//! TWO IMPEDANCE MISMATCHES, HANDLED HERE RATHER THAN IN EITHER ENGINE
//!
//! 1. The seam's `DOCUMENT` is node `0`. In htmlbox's arena, slot 0 is the
//!    SENTINEL meaning "no node" — `dom_append_child(0, child)` returns early.
//!    Left alone, appending to the document would silently do nothing. `to_hb`
//!    and `from_hb` below translate between the two spellings.
//!
//! 2. htmlbox dispatches events to callbacks synchronously; the seam pulls
//!    them with `DrainEvents`. A queue fed by htmlbox's global form-event
//!    callback bridges the two, so neither engine changes shape.

use std::collections::{HashMap, VecDeque};
use std::sync::{Arc, Mutex, OnceLock};

use rhtmledit::types::{Document, FormEvent, FormEventKind};

use crate::engine::{
    DOCUMENT, DocumentId, DomOp, DomValue, EventOp, EventValue, NodeId, ScheduleOp, ScheduleValue,
    WebEngine, WindowOp, WindowValue,
};

/// The viewport a document is laid out against before anyone resizes it.
/// `WindowOp::ResizeTo` is what changes it afterwards.
///
/// The size `vybe_widgets` gives a form it was told nothing about
/// (`dom.rs:735`). A default is a user-agent choice and either number would be
/// defensible on its own — but two engines that answer differently make every
/// program that declares no size render at two sizes, which turns a swap into a
/// diff nobody can read.
const DEFAULT_VIEWPORT_W: f32 = 800.0;
const DEFAULT_VIEWPORT_H: f32 = 600.0;

/// One document plus the queue its events land in.
struct Entry {
    doc: Document,
    /// Shared with the form-event callback installed on `doc`, which is why
    /// this is an `Arc` and not a plain field: the callback outlives the
    /// borrow that registered it.
    events: Arc<Mutex<VecDeque<(NodeId, String)>>>,
}

#[derive(Default)]
struct Docs {
    next: DocumentId,
    entries: HashMap<DocumentId, Mutex<Entry>>,
}

/// `Document` is `Send` (htmlbox declares it for its parallel cascade) but not
/// `Sync` — a `Mutex` around it is both, which is what the trait requires.
/// Same shape `vybe_widgets::dom` uses for its own document table.
fn docs() -> &'static Mutex<Docs> {
    static DOCS: OnceLock<Mutex<Docs>> = OnceLock::new();
    DOCS.get_or_init(|| Mutex::new(Docs::default()))
}

/// Borrow a document. Mirrors `engine_widgets::with_document` so the window
/// runner can reach a form the same way under either engine.
pub fn with_document<T>(id: DocumentId, f: impl FnOnce(&mut Document) -> T) -> Option<T> {
    let map = docs().lock().ok()?;
    let entry = map.entries.get(&id)?;
    let mut entry = entry.lock().ok()?;
    Some(f(&mut entry.doc))
}

fn with_entry<T>(id: DocumentId, f: impl FnOnce(&mut Entry) -> T) -> Option<T> {
    let map = docs().lock().ok()?;
    let entry = map.entries.get(&id)?;
    let mut entry = entry.lock().ok()?;
    Some(f(&mut entry))
}

// ── Node id translation ─────────────────────────────────────────────────────

/// Seam node → htmlbox node. `DOCUMENT` (0) is the document itself, which in
/// htmlbox is the root box; 0 in the arena means "no node".
///
/// For ADDRESSING a node this is right — htmlbox has no separate document node,
/// so the document's stand-in is `<html>`. For INSERTING one it is not: see
/// [`insertion_parent`].
fn to_hb(doc: &Document, node: NodeId) -> u32 {
    if node == DOCUMENT { doc.root.node_id } else { node as u32 }
}

/// The node a DOCUMENT-addressed CONTENT operation means: the **body**.
///
/// Style is the case that matters. A .NET form IS the document, so `Form.Width`
/// and `Form.BackColor` lower to `setStyleProperty` on it — and a page's own box
/// and background are the BODY's, not the document's. Sent to `<html>` they
/// styled a box the renderer takes its canvas colour and its page size from
/// somewhere else entirely, so a form came out white and 1024x768 whatever it
/// asked for. `vybe_widgets` keeps these as the document's own declarations and
/// applies them to the form, which is the same box by another name.
///
/// Falls back to the root when there is no body — an XML document has none, and
/// answering `0` there would mean "no node" and drop the write silently.
fn content_node(doc: &Document, node: NodeId) -> u32 {
    if node != DOCUMENT {
        return node as u32;
    }
    doc.body().unwrap_or_else(|| to_hb(doc, DOCUMENT))
}

/// Where a node addressed to the DOCUMENT actually goes: the **body**.
///
/// The same rule `vybe_widgets::dom::Document::content_parent` states, because
/// it is the DOM's rule rather than either engine's: a Document takes exactly
/// one element child, so `document.appendChild(<p>)` is a
/// `HierarchyRequestError` in a browser. A caller that says "the document"
/// means the body. `<html>`/`<head>`/`<body>` ARE the document's structure, so
/// a caller that spells one out is obeyed where it put it.
///
/// Without this htmlbox hung every control off `<html>`, a sibling of `<head>`
/// and `<body>` rather than a child of either. The tree laid out — that is what
/// the timings reported — and painted nothing a page would recognise as
/// content, because none of it was in the body.
fn insertion_parent(doc: &Document, parent: NodeId, child: NodeId) -> u32 {
    if parent != DOCUMENT {
        return parent as u32;
    }
    let structural = matches!(
        doc.local_name(child as u32).as_str(),
        "html" | "head" | "body"
    );
    match (structural, doc.body()) {
        (false, Some(body)) => body,
        _ => to_hb(doc, parent),
    }
}

/// htmlbox node → seam node. The root box answers as `DOCUMENT` so that
/// walking up from `<body>` lands on the document, as the DOM says it should.
fn from_hb(doc: &Document, id: u32) -> NodeId {
    if id == doc.root.node_id || id == 0 { DOCUMENT } else { id as NodeId }
}

/// The DOM event name for a form interaction. htmlbox reports what HAPPENED;
/// the seam is keyed by the name a listener was registered under.
fn event_names(kind: &FormEventKind) -> Vec<&'static str> {
    match kind {
        FormEventKind::Input(_) => vec!["input"],
        FormEventKind::Change(_) => vec!["change"],
        // Ticking a checkbox is a click AND a change in HTML, and a listener
        // may be registered under either. Both are queued.
        FormEventKind::Toggle(_) => vec!["click", "change"],
        FormEventKind::Click(_) => vec!["click"],
        FormEventKind::Submit(_) => vec!["submit"],
        FormEventKind::Focus => vec!["focus"],
        _ => vec![],
    }
}

struct HtmlBox;

impl WebEngine for HtmlBox {
    fn new_document(&self, title: &str) -> DocumentId {
        // Parsed rather than `Document::new()`: the seam expects a real
        // `<head>`/`<body>` skeleton, and `<title>` is where `DomOp::Title`
        // reads from.
        let escaped = rhtmledit::html::serializer::escape_html(title);
        // **The app shell, said in CSS.** A document opened by a program is a
        // window's worth of page, not a scrolling article, and the two rules
        // below are what every app stylesheet on the web opens with.
        //
        // `height: 100%` because a percentage height resolves against the
        // parent's, and a page's root boxes are auto by default: a Flutter
        // `Scaffold` is `height: 100%`, so with nothing above it every
        // `Expanded` row underneath resolved to zero and the whole app
        // collapsed to a line of text. `vybe_widgets` says the same thing by
        // making its root boxes fill the viewport (`fit_body_to_viewport`),
        // which is the same fact in the toolkit's vocabulary.
        //
        // `margin: 0` because the UA's `body { margin: 8px }` is for prose. A
        // form that asks for exactly the window's width overflows it by those
        // 16px and gets a scrollbar it never asked for.
        //
        // An AUTHOR sheet, so a page that wants the margin back can simply set
        // it — this is the shell's own styling, not a change to what htmlbox
        // believes about HTML.
        let html = format!(
            "<html><head><title>{escaped}</title>\
             <style>html, body {{ height: 100%; margin: 0; }}</style>\
             </head><body></body></html>"
        );
        let doc = rhtmledit::load_html(&html, DEFAULT_VIEWPORT_W);

        let events: Arc<Mutex<VecDeque<(NodeId, String)>>> =
            Arc::new(Mutex::new(VecDeque::new()));

        let mut entry = Entry { doc, events: Arc::clone(&events) };

        // The bridge: htmlbox calls this synchronously as interactions happen,
        // and `DrainEvents` pulls whatever accumulated since the last drain.
        let sink = Arc::clone(&events);
        let root_id = entry.doc.root.node_id;
        entry.doc.on_form_event = Some(Box::new(move |e: &FormEvent| {
            if let Ok(mut q) = sink.lock() {
                let node = if e.element == root_id || e.element == 0 {
                    DOCUMENT
                } else {
                    e.element as NodeId
                };
                for name in event_names(&e.kind) {
                    q.push_back((node, name.to_string()));
                }
            }
        }));

        let mut map = match docs().lock() {
            Ok(m) => m,
            Err(_) => return 0,
        };
        map.next += 1;
        let id = map.next;
        map.entries.insert(id, Mutex::new(entry));
        id
    }

    fn new_xml_document(&self, title: &str) -> DocumentId {
        // Built by the HTML parser, then marked XML — the skeleton is the same
        // tree either way, and what the kind changes is NAME FOLDING: an XML
        // document is case-sensitive, so `<Rect>` and `<rect>` stay distinct
        // instead of both folding to `rect`.
        //
        // XML TEXT is not parsed here and does not need to be: `dom_parser.rs`
        // owns `DOMParser`/`XMLSerializer` above the seam and builds trees
        // through `CreateElementNS` and friends, so the tokenizer is shared by
        // both engines rather than duplicated inside either.
        let id = self.new_document(title);
        with_document(id, |doc| {
            doc.kind = rhtmledit::types::DocumentKind::Xml;
        });
        id
    }

    fn document(&self, document: DocumentId, op: DomOp) -> DomValue {
        with_entry(document, |entry| {
            let doc = &mut entry.doc;
            match op {
                // ── Creation ──
                DomOp::CreateElement { tag, input_type } => {
                    let id = doc.create_element(&tag);
                    if !input_type.is_empty() {
                        doc.set_attribute(id, "type", &input_type);
                    }
                    DomValue::Node(from_hb(doc, id))
                }
                DomOp::CreateTextNode(data) => {
                    DomValue::Node(doc.create_text_node(&data) as NodeId)
                }
                DomOp::CreateComment(data) => {
                    DomValue::Node(doc.create_comment(&data) as NodeId)
                }

                // ── XML ──
                DomOp::CreateElementNS { namespace, qualified_name, input_type } => {
                    let id = doc.create_element_ns(&namespace, &qualified_name);
                    if !input_type.is_empty() {
                        doc.set_attribute(id, "type", &input_type);
                    }
                    DomValue::Node(id as NodeId)
                }
                DomOp::CreateCDataSection(data) => {
                    DomValue::Node(doc.create_cdata_section(&data) as NodeId)
                }
                DomOp::CreateProcessingInstruction { target, data } => {
                    DomValue::Node(doc.create_processing_instruction(&target, &data) as NodeId)
                }
                DomOp::NamespaceUri(n) => match doc.namespace_uri(to_hb(doc, n)) {
                    Some(uri) => DomValue::Text(uri),
                    None => DomValue::Null,
                },
                DomOp::Prefix(n) => match doc.prefix(to_hb(doc, n)) {
                    Some(prefix) => DomValue::Text(prefix),
                    None => DomValue::Null,
                },
                DomOp::LocalName(n) => DomValue::Text(doc.local_name(to_hb(doc, n))),
                DomOp::SetAttributeNS { node, namespace, qualified_name, value } => {
                    doc.set_attribute_ns(to_hb(doc, node), &namespace, &qualified_name, &value);
                    DomValue::None
                }
                DomOp::GetAttributeNS { node, namespace, local_name } => {
                    match doc.get_attribute_ns(to_hb(doc, node), &namespace, &local_name) {
                        Some(v) => DomValue::Text(v),
                        None => DomValue::Null,
                    }
                }

                // ── Queries ──
                DomOp::GetElementById(id) => match doc.get_element_by_id(&id) {
                    Some(n) => DomValue::Node(from_hb(doc, n)),
                    None => DomValue::Null,
                },
                DomOp::ElementsByTag(tag) => {
                    let nodes = doc.get_elements_by_tag_name(&tag);
                    DomValue::Nodes(nodes.iter().map(|n| from_hb(doc, *n)).collect())
                }
                DomOp::QuerySelector(sel) => match doc.query_selector(&sel) {
                    Some(n) => DomValue::Node(from_hb(doc, n)),
                    None => DomValue::Null,
                },
                DomOp::QuerySelectorAll(sel) => {
                    let nodes = doc.query_selector_all(&sel);
                    DomValue::Nodes(nodes.iter().map(|n| from_hb(doc, *n)).collect())
                }
                DomOp::Title => DomValue::Text(doc.title()),
                DomOp::SetTitle(t) => {
                    doc.set_title(&t);
                    DomValue::None
                }

                // ── Tree ──
                DomOp::AppendChild { parent, child } => {
                    let p = insertion_parent(doc, parent, child);
                    doc.append_child(p, child as u32);
                    DomValue::Bool(true)
                }
                // htmlbox unlinks a child from whatever parent it is ACTUALLY
                // under, so there is no parent to redirect here — which is what
                // keeps the redirect on the way in from leaving a node linked
                // into the body and unlinked from the document.
                DomOp::RemoveChild { child, .. } => {
                    doc.remove_child(child as u32);
                    DomValue::Bool(true)
                }
                DomOp::InsertBefore { parent, child, reference } => {
                    let p = insertion_parent(doc, parent, child);
                    doc.insert_before(p, child as u32, reference as u32);
                    DomValue::Bool(true)
                }
                DomOp::ReplaceChild { parent, new_child, old_child } => {
                    let p = insertion_parent(doc, parent, new_child);
                    DomValue::Bool(doc.replace_child(p, new_child as u32, old_child as u32))
                }
                DomOp::CloneNode { node, deep } => {
                    let clone = doc.clone_node(to_hb(doc, node), deep);
                    if clone == 0 { DomValue::Null } else { DomValue::Node(clone as NodeId) }
                }
                // The document is not its own element. htmlbox has no node for
                // it, so `to_hb` answers `<html>` — right for reaching into the
                // tree, wrong for the two questions that ask what the node IS.
                // DOM §4.4: `9` and `#document`, which is what `vybe_widgets`
                // answers, and the seam is one API or it is not one.
                DomOp::NodeType(DOCUMENT) => DomValue::Number(9.0),
                DomOp::NodeName(DOCUMENT) => DomValue::Text("#document".to_string()),
                DomOp::NodeType(n) => {
                    DomValue::Number(f64::from(doc.node_type(to_hb(doc, n))))
                }
                DomOp::NodeName(n) => DomValue::Text(doc.node_name(to_hb(doc, n))),
                DomOp::NodeValue(n) => match doc.node_value(to_hb(doc, n)) {
                    Some(v) => DomValue::Text(v),
                    None => DomValue::Null,
                },
                DomOp::ParentNode(n) => {
                    let parent = doc.parent_node(to_hb(doc, n));
                    if parent == 0 { DomValue::Null } else { DomValue::Node(from_hb(doc, parent)) }
                }
                DomOp::ChildNodes(n) => {
                    let kids = doc.child_nodes(to_hb(doc, n));
                    DomValue::Nodes(kids.iter().map(|c| from_hb(doc, *c)).collect())
                }
                DomOp::IsConnected(n) => DomValue::Bool(doc.is_connected(to_hb(doc, n))),
                DomOp::InnerHtml(n) => DomValue::Text(doc.inner_html(to_hb(doc, n))),
                // Parsing re-enters `apply` to build the tree, which would be a
                // second borrow of the document held here. Dispatched before
                // the lock instead — same split `engine_widgets` makes.
                DomOp::SetInnerHtml { .. } => DomValue::None,
                DomOp::OuterHtml(n) => DomValue::Text(doc.outer_html(to_hb(doc, n))),
                // Both markup setters are dispatched before the lock, for the
                // same reason the `innerHTML` one is.
                DomOp::SetOuterHtml { .. } => DomValue::None,
                DomOp::InsertAdjacentHtml { .. } => DomValue::None,
                DomOp::CreateDocumentFragment => {
                    DomValue::Node(doc.create_document_fragment() as NodeId)
                }
                // Cross-document, so dispatched before the lock like the
                // markup setters — see `apply`.
                DomOp::ImportNode { .. } => DomValue::Null,
                DomOp::TextContent(n) => DomValue::Text(doc.text_content(to_hb(doc, n))),
                // **A document's text is its TITLE, and writing it must not
                // touch the tree.** `textContent` replaces all of a node's
                // children with one text node — right for an element, and for
                // the document it means `<head>` and `<body>` are DELETED.
                //
                // That is what emptied every .NET form under this engine: a
                // form's caption is `Form.Text`, the form IS the document, so
                // `Form.Text = "…"` wiped the body and every control appended
                // afterwards hung off a bodyless `<html>` and rendered nothing.
                // DOM §4.4 says `Document.textContent` is null and setting it
                // does nothing; `vybe_widgets` answers the title, and one seam
                // cannot have two answers.
                DomOp::TextContent(DOCUMENT) => DomValue::Text(doc.title()),
                DomOp::SetTextContent(DOCUMENT, t) => {
                    doc.set_title(&t);
                    DomValue::None
                }
                DomOp::SetTextContent(n, t) => {
                    doc.set_text_content(to_hb(doc, n), &t);
                    DomValue::None
                }

                // ── Attributes ──
                DomOp::SetAttribute(n, name, value) => {
                    doc.set_attribute(to_hb(doc, n), &name, &value);
                    DomValue::None
                }
                DomOp::GetAttribute(n, name) => {
                    match doc.get_attribute(to_hb(doc, n), &name) {
                        Some(v) => DomValue::Text(v),
                        None => DomValue::Null,
                    }
                }
                DomOp::AttributeNames(n) => {
                    DomValue::Texts(doc.get_attribute_names(to_hb(doc, n)))
                }
                DomOp::RemoveAttribute(n, name) => {
                    doc.remove_attribute(to_hb(doc, n), &name);
                    DomValue::None
                }
                DomOp::SetStyleProperty(n, p, v) => {
                    doc.set_style_property(content_node(doc, n), &p, &v);
                    DomValue::None
                }
                // The DECLARED value — what was authored, un-resolved. htmlbox
                // already answered this way, which is why it disagreed with the
                // old `vybe_widgets` for `left`/`top`/`width`/`height`.
                DomOp::GetStyleProperty(n, p) => {
                    DomValue::Text(
                        doc.get_style_property(content_node(doc, n), &p).unwrap_or_default(),
                    )
                }
                // The RESOLVED value. Geometry comes off the laid-out rect;
                // everything else falls back to the declared value, matching
                // the floor `vybe_widgets` sets.
                DomOp::ComputedStyleProperty(n, p) => {
                    let node = content_node(doc, n);
                    DomValue::Text(doc.computed_style_property(node, &p))
                }

                // ── Form controls ──
                //
                // `checked` and `value` are CONTENT ATTRIBUTES in htmlbox, which
                // is the HTML model, so these are attribute reads rather than a
                // parallel control-state store. Interaction writes them on the
                // render tree and `sync_form_state_to_arena` reconciles, so a
                // read here sees the user's last click.
                // NOT a plain `value` attribute read: `value` means the text
                // content of a `<textarea>` and the selected option's value on
                // a `<select>`. See `Document::dom_value`.
                DomOp::Value(n) => DomValue::Text(doc.value(to_hb(doc, n))),
                DomOp::SetValue(n, v) => {
                    doc.set_value(to_hb(doc, n), &v);
                    DomValue::None
                }

                // ── HTMLSelectElement: the items are `<option>` children ──
                DomOp::AddItem(n, t) => {
                    doc.add_item(to_hb(doc, n), &t);
                    DomValue::None
                }
                DomOp::RemoveItem(n, i) => {
                    doc.remove_item(to_hb(doc, n), i);
                    DomValue::None
                }
                DomOp::ClearItems(n) => {
                    doc.clear_items(to_hb(doc, n));
                    DomValue::None
                }
                DomOp::ItemText(n, i) => DomValue::Text(doc.item_text(to_hb(doc, n), i)),
                DomOp::SetItemText(n, i, t) => {
                    doc.set_item_text(to_hb(doc, n), i, &t);
                    DomValue::None
                }
                DomOp::SelectedIndex(n) => {
                    DomValue::Number(f64::from(doc.selected_index(to_hb(doc, n))))
                }
                DomOp::SetSelectedIndex(n, i) => {
                    doc.set_selected_index(to_hb(doc, n), i);
                    DomValue::None
                }
                DomOp::Checked(n) => DomValue::Bool(doc.checked(to_hb(doc, n))),
                DomOp::SetChecked(n, c) => {
                    doc.set_checked(to_hb(doc, n), c);
                    DomValue::None
                }

                DomOp::Focus(n) => {
                    doc.focus(to_hb(doc, n));
                    DomValue::None
                }

                // ── Events ──
                // htmlbox hit-tests, dispatches the DOM event and — for a click
                // on a control — calls `on_form_event`, which is the callback
                // this file installed at `new_document`. So the input arrives
                // here and comes back out of `DrainEvents` with no further
                // wiring: the bridge was already built, nothing was crossing it.
                DomOp::DispatchPointer {
                    kind,
                    client_x,
                    client_y,
                    button,
                } => {
                    use rhtmledit::dom::HtmlEventType;
                    let etype = match kind.as_str() {
                        "mousedown" => HtmlEventType::MouseDown,
                        "mouseup" => HtmlEventType::MouseUp,
                        _ => HtmlEventType::MouseMove,
                    };
                    // `MouseEvent.button` is signed and `process_mouse_event`
                    // takes the same three values unsigned; anything else is
                    // not a button htmlbox knows and is treated as the primary
                    // one, which is what it does with an unrecognised device.
                    let button = u8::try_from(button).unwrap_or(0);
                    DomValue::Bool(doc.process_mouse_event(etype, (client_x, client_y), button))
                }
                DomOp::DrainEvents => {
                    let drained: Vec<(NodeId, String)> = match entry.events.lock() {
                        Ok(mut q) => q.drain(..).collect(),
                        Err(_) => Vec::new(),
                    };
                    DomValue::Events(drained)
                }

                // ── HTMLDialogElement ──
                DomOp::ShowDialog { node, modal } => {
                    doc.show_dialog(to_hb(doc, node), modal);
                    DomValue::None
                }
                DomOp::CloseDialog(node) => {
                    doc.close_dialog(to_hb(doc, node));
                    DomValue::None
                }
                DomOp::DialogOpen(node) => {
                    DomValue::Bool(doc.dialog_open(to_hb(doc, node)))
                }

                DomOp::BoundingClientRect(node) => {
                    match doc.get_bounding_client_rect(to_hb(doc, node)) {
                        Some(rect) => DomValue::Rect {
                            x: rect.x as f64,
                            y: rect.y as f64,
                            width: rect.w as f64,
                            height: rect.h as f64,
                        },
                        None => DomValue::Rect { x: 0.0, y: 0.0, width: 0.0, height: 0.0 },
                    }
                }
                DomOp::CanvasSize(node) => match doc.get_bounding_client_rect(to_hb(doc, node)) {
                    Some(r) => DomValue::Pair(f64::from(r.w), f64::from(r.h)),
                    None => DomValue::None,
                },

                // ── Not yet covered by this engine ──
                //
                // Each of these is a GAP, listed rather than quietly answered:
                //
                //   select/option — `SelectedIndex`, `SetSelectedIndex`,
                //     `ItemText`, `SetItemText`, `AddItem`, `RemoveItem`,
                //     `ClearItems`. htmlbox has `widgets/select.rs` with
                //     `select_index`/`selected_text`, but no DOM spelling over
                //     `<option>` children yet.
                //   `ShowPicker` — needs the UA's own file/colour chooser.
                //   XML — the namespace, PI and CDATA ops. `NodeType` has no
                //     variant for them, so this is a missing MODEL, not a
                //     missing wrapper.
                _ => DomValue::Null,
            }
        })
        .unwrap_or(DomValue::None)
    }

    fn window(&self, op: WindowOp) -> WindowValue {
        match op {
            // The id-minting ops stay HERE so there is one document table.
            WindowOp::Open { target, .. } => WindowValue::Window(self.new_document(&target)),
            // A window and its document are the same handle under this engine:
            // htmlbox has no separate window object, and a `Document` IS the
            // browsing context a tab renders (which is what `browser.rs` does
            // with one `Document` per tab).
            WindowOp::Document(w) => WindowValue::Document(w),
            WindowOp::DefaultView(d) => WindowValue::Window(d),
            WindowOp::AdoptTopLevel(d) => WindowValue::Window(d),
            WindowOp::Closed(w) => WindowValue::Bool(
                docs().lock().map(|m| !m.entries.contains_key(&w)).unwrap_or(true),
            ),
            WindowOp::Close(w) => {
                if let Ok(mut m) = docs().lock() {
                    m.entries.remove(&w);
                }
                WindowValue::None
            }
            // **The page's own box, when it declares one.** A .NET form sizes
            // itself by writing `width`/`height` on the document, which is the
            // body — so a form that asked for 800x600 got a 1024x768 window and
            // a screenshot padded with 224 columns of nothing.
            //
            // A page that declares no size keeps the default: that is a window
            // the user agent chose, and it is what an ordinary HTML page gets.
            WindowOp::InnerSize(w) => {
                let declared = with_document(w, |doc| {
                    let body = content_node(doc, DOCUMENT);
                    let axis = |property: &str| {
                        doc.get_style_property(body, property)
                            .and_then(|value| value.trim().strip_suffix("px")?.parse::<f32>().ok())
                            .filter(|px| *px >= 1.0)
                    };
                    (axis("width"), axis("height"))
                });
                match declared {
                    Some((Some(width), Some(height))) => {
                        WindowValue::Pair(width as f64, height as f64)
                    }
                    _ => WindowValue::Pair(DEFAULT_VIEWPORT_W as f64, DEFAULT_VIEWPORT_H as f64),
                }
            }
            // Geometry, chrome and the message boxes are host concerns, not
            // DOM — identical under either engine, so they delegate.
            other => crate::engine_widgets::Widgets.window(other),
        }
    }

    fn events(&self, op: EventOp) -> EventValue {
        // All four delegate. This is the RAW INPUT queue — process-wide, filled
        // by the winit shell — not the per-document DOM queue that
        // `DomOp::DrainEvents` serves. Nothing about it changes with the DOM
        // engine, so re-implementing it here would fork input delivery for no
        // gain.
        crate::engine_widgets::Widgets.events(op)
    }

    fn schedule(&self, op: ScheduleOp) -> ScheduleValue {
        // Timers and frames are wall-clock bookkeeping with no DOM in them.
        crate::engine_widgets::Widgets.schedule(op)
    }
}

/// Install htmlbox as the web engine.
pub fn install() {
    crate::engine::set_engine(Arc::new(HtmlBox));
}
