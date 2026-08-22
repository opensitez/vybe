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

/// The width a document is laid out against before anyone resizes it.
/// `WindowOp::ResizeTo` is what changes it afterwards.
const DEFAULT_VIEWPORT_W: f32 = 1024.0;
const DEFAULT_VIEWPORT_H: f32 = 768.0;

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
fn to_hb(doc: &Document, node: NodeId) -> u32 {
    if node == DOCUMENT { doc.root.node_id } else { node as u32 }
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
        let html = format!(
            "<html><head><title>{escaped}</title></head><body></body></html>"
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
                    let p = to_hb(doc, parent);
                    doc.append_child(p, child as u32);
                    DomValue::Bool(true)
                }
                DomOp::RemoveChild { child, .. } => {
                    doc.remove_child(child as u32);
                    DomValue::Bool(true)
                }
                DomOp::InsertBefore { parent, child, reference } => {
                    let p = to_hb(doc, parent);
                    doc.insert_before(p, child as u32, reference as u32);
                    DomValue::Bool(true)
                }
                DomOp::ReplaceChild { parent, new_child, old_child } => {
                    let p = to_hb(doc, parent);
                    DomValue::Bool(doc.replace_child(p, new_child as u32, old_child as u32))
                }
                DomOp::CloneNode { node, deep } => {
                    let clone = doc.clone_node(to_hb(doc, node), deep);
                    if clone == 0 { DomValue::Null } else { DomValue::Node(clone as NodeId) }
                }
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
                    doc.set_style_property(to_hb(doc, n), &p, &v);
                    DomValue::None
                }
                // The DECLARED value — what was authored, un-resolved. htmlbox
                // already answered this way, which is why it disagreed with the
                // old `vybe_widgets` for `left`/`top`/`width`/`height`.
                DomOp::GetStyleProperty(n, p) => {
                    DomValue::Text(doc.get_style_property(to_hb(doc, n), &p).unwrap_or_default())
                }
                // The RESOLVED value. Geometry comes off the laid-out rect;
                // everything else falls back to the declared value, matching
                // the floor `vybe_widgets` sets.
                DomOp::ComputedStyleProperty(n, p) => {
                    let node = to_hb(doc, n);
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
            WindowOp::InnerSize(_) => {
                WindowValue::Pair(DEFAULT_VIEWPORT_W as f64, DEFAULT_VIEWPORT_H as f64)
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
