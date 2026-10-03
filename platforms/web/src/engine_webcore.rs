//! webcore as the engine behind `web:*`.
//!
//! Implements the shared `engine.rs` contract for the in-process browser.
//!
//! WHAT THIS FILE OWNS AND WHAT IT DELEGATES
//!
//! `DocumentId` is ONE namespace shared by `document()` and `window()`:
//! `WindowOp::Open`, `Document`, `DefaultView` and `AdoptTopLevel` all mint or
//! return one. WebCore owns the whole browsing context: all of `DomOp`
//! and the id-minting `WindowOp`s.
//!
//! `DomOp::DrainEvents` is per-document listener dispatch. `EventOp` is raw
//! browser input, held by webcore's own queue. Scheduling also belongs to the
//! selected browser, never to the other engine.
//!
//! TWO IMPEDANCE MISMATCHES, HANDLED HERE RATHER THAN IN EITHER ENGINE
//!
//! 1. The seam's `DOCUMENT` is node `0`. In webcore's arena, slot 0 is the
//!    SENTINEL meaning "no node" — `dom_append_child(0, child)` returns early.
//!    Left alone, appending to the document would silently do nothing. `to_hb`
//!    and `from_hb` below translate between the two spellings.
//!
//! 2. webcore dispatches events to callbacks synchronously; the seam pulls
//!    them with `DrainEvents`. DOM listeners and form callbacks feed the queue,
//!    so neither engine changes shape.

use std::collections::{HashMap, VecDeque};
use std::sync::{Arc, Mutex, OnceLock};

use webcore::types::{Document, FormEvent, FormEventKind};

use crate::engine::{
    DOCUMENT, DocumentId, DomEventRecord, DomOp, DomValue, EventOp, EventValue, NodeId, PickerOp,
    ScheduleOp, ScheduleValue, UiEventFields, WebEngine, WindowOp, WindowValue,
};

fn into_webcore_event(f: UiEventFields) -> webcore::ui_events::UiEvent {
    webcore::ui_events::UiEvent {
        kind: f.kind, key: f.key, code: f.code, key_code: f.key_code,
        client_x: f.client_x, client_y: f.client_y, button: f.button,
        buttons: f.buttons, delta_y: f.delta_y, ctrl_key: f.ctrl_key,
        shift_key: f.shift_key, alt_key: f.alt_key, meta_key: f.meta_key,
        repeat: f.repeat,
    }
}

fn from_webcore_event(f: webcore::ui_events::UiEvent) -> UiEventFields {
    UiEventFields {
        kind: f.kind, key: f.key, code: f.code, key_code: f.key_code,
        client_x: f.client_x, client_y: f.client_y, button: f.button,
        buttons: f.buttons, delta_y: f.delta_y, ctrl_key: f.ctrl_key,
        shift_key: f.shift_key, alt_key: f.alt_key, meta_key: f.meta_key,
        repeat: f.repeat,
    }
}

trait WebcoreDocumentLayoutCompat {
    fn flush_layout(&mut self);
}

impl WebcoreDocumentLayoutCompat for Document {
    fn flush_layout(&mut self) {
        let width = self.root.layout.last_containing_width;
        if width <= 0.0 {
            return;
        }
        webcore::LayoutEngine::new().layout(self, width);
    }
}

/// The viewport a document is laid out against before anyone resizes it.
/// `WindowOp::ResizeTo` is what changes it afterwards.
///
/// The size `widgets` gives a form it was told nothing about
/// (`dom.rs:735`). A default is a user-agent choice and either number would be
/// defensible on its own — but two engines that answer differently make every
/// program that declares no size render at two sizes, which turns a swap into a
/// diff nobody can read.
const DEFAULT_VIEWPORT_W: f32 = 800.0;
const DEFAULT_VIEWPORT_H: f32 = 600.0;

/// One document plus the queue its events land in.
struct Entry {
    doc: Document,
    resources_dirty: bool,
    /// Shared with the form-event callback installed on `doc`, which is why
    /// this is an `Arc` and not a plain field: the callback outlives the
    /// borrow that registered it.
    events: Arc<Mutex<VecDeque<(NodeId, String)>>>,
    targeted_events: Arc<Mutex<VecDeque<DomEventRecord>>>,
    observed: HashMap<(NodeId, String), u32>,
}

#[derive(Default)]
struct Docs {
    next: DocumentId,
    entries: HashMap<DocumentId, Mutex<Entry>>,
}

/// `Document` is `Send` (webcore declares it for its parallel cascade) but not
/// `Sync` — a `Mutex` around it is both, which is what the trait requires.
fn docs() -> &'static Mutex<Docs> {
    static DOCS: OnceLock<Mutex<Docs>> = OnceLock::new();
    DOCS.get_or_init(|| Mutex::new(Docs::default()))
}

/// Borrow a WebCore document for in-process presentation.
pub fn with_document<T>(id: DocumentId, f: impl FnOnce(&mut Document) -> T) -> Option<T> {
    let map = docs().lock().ok()?;
    let entry = map.entries.get(&id)?;
    let mut entry = entry.lock().ok()?;
    Some(f(&mut entry.doc))
}

pub fn with_document_resources<T>(
    id: DocumentId,
    f: impl FnOnce(&mut Document, &mut bool) -> T,
) -> Option<T> {
    with_entry(id, |entry| f(&mut entry.doc, &mut entry.resources_dirty))
}

fn with_entry<T>(id: DocumentId, f: impl FnOnce(&mut Entry) -> T) -> Option<T> {
    let map = docs().lock().ok()?;
    let entry = map.entries.get(&id)?;
    let mut entry = entry.lock().ok()?;
    Some(f(&mut entry))
}

// ── Node id translation ─────────────────────────────────────────────────────

/// Seam node → webcore node. `DOCUMENT` (0) is the document itself, which in
/// webcore is the root box; 0 in the arena means "no node".
///
/// For ADDRESSING a node this is right — webcore has no separate document node,
/// so the document's stand-in is `<html>`. For INSERTING one it is not: see
/// [`insertion_parent`].
fn to_hb(doc: &Document, node: NodeId) -> u32 {
    if node == DOCUMENT {
        doc.root.node_id
    } else {
        node as u32
    }
}

/// The node a DOCUMENT-addressed CONTENT operation means: the **body**.
///
/// Style is the case that matters. A .NET form IS the document, so `Form.Width`
/// and `Form.BackColor` lower to `setStyleProperty` on it — and a page's own box
/// and background are the BODY's, not the document's. Sent to `<html>` they
/// styled a box the renderer takes its canvas colour and its page size from
/// somewhere else entirely, so a form came out white and 1024x768 whatever it
/// asked for. `widgets` keeps these as the document's own declarations and
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
/// A Document takes exactly
/// one element child, so `document.appendChild(<p>)` is a
/// `HierarchyRequestError` in a browser. A caller that says "the document"
/// means the body. `<html>`/`<head>`/`<body>` ARE the document's structure, so
/// a caller that spells one out is obeyed where it put it.
///
/// Without this webcore hung every control off `<html>`, a sibling of `<head>`
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

/// webcore node → seam node. The root box answers as `DOCUMENT` so that
/// walking up from `<body>` lands on the document, as the DOM says it should.
fn from_hb(doc: &Document, id: u32) -> NodeId {
    if id == doc.root.node_id || id == 0 {
        DOCUMENT
    } else {
        id as NodeId
    }
}

/// The DOM event name for a form interaction. webcore reports what HAPPENED;
/// the seam is keyed by the name a listener was registered under.
fn event_names(kind: &FormEventKind) -> Vec<&'static str> {
    match kind {
        FormEventKind::Input(_) => vec!["input"],
        FormEventKind::Change(_) => vec!["change"],
        FormEventKind::Toggle(_) => vec!["click", "change"],
        FormEventKind::Click(_) => vec!["click"],
        FormEventKind::Submit(_) => vec!["submit"],
        FormEventKind::Focus => vec!["focus"],
        _ => vec![],
    }
}

struct WebCore;

impl WebEngine for WebCore {
    fn new_document(&self, title: &str) -> DocumentId {
        // Parsed rather than `Document::new()`: the seam expects a real
        // `<head>`/`<body>` skeleton, and `<title>` is where `DomOp::Title`
        // reads from.
        let escaped = webcore::html::serializer::escape_html(title);
        // **The app shell, said in CSS.** A document opened by a program is a
        // window's worth of page, not a scrolling article, and the two rules
        // below are what every app stylesheet on the web opens with.
        //
        // `height: 100%` because a percentage height resolves against the
        // parent's, and a page's root boxes are auto by default: a Flutter
        // `Scaffold` is `height: 100%`, so with nothing above it every
        // `Expanded` row underneath resolved to zero and the whole app
        // collapsed to a line of text. `widgets` says the same thing by
        // making its root boxes fill the viewport (`fit_body_to_viewport`),
        // which is the same fact in the toolkit's vocabulary.
        //
        // `margin: 0` because the UA's `body { margin: 8px }` is for prose. A
        // form that asks for exactly the window's width overflows it by those
        // 16px and gets a scrollbar it never asked for.
        //
        // An AUTHOR sheet, so a page that wants the margin back can simply set
        // it — this is the shell's own styling, not a change to what webcore
        // believes about HTML.
        let html = format!(
            "<html><head><title>{escaped}</title>\
             <style>html, body {{ height: 100%; margin: 0; }}</style>\
             </head><body></body></html>"
        );
        let mut doc = webcore::load_html(&html, DEFAULT_VIEWPORT_W);
        doc.set_viewport(DEFAULT_VIEWPORT_W, DEFAULT_VIEWPORT_H);

        let events: Arc<Mutex<VecDeque<(NodeId, String)>>> = Arc::new(Mutex::new(VecDeque::new()));

        let mut entry = Entry {
            doc,
            resources_dirty: true,
            events: Arc::clone(&events),
            targeted_events: Arc::new(Mutex::new(VecDeque::new())),
            observed: HashMap::new(),
        };

        // The bridge: webcore calls this synchronously as interactions happen,
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
            doc.kind = webcore::types::DocumentKind::Xml;
        });
        id
    }

    fn document(&self, document: DocumentId, op: DomOp) -> DomValue {
        with_entry(document, |entry| {
            if matches!(
                &op,
                DomOp::AppendChild { .. }
                    | DomOp::RemoveChild { .. }
                    | DomOp::InsertBefore { .. }
                    | DomOp::ReplaceChild { .. }
                    | DomOp::SetInnerHtml { .. }
                    | DomOp::SetOuterHtml { .. }
                    | DomOp::InsertAdjacentHtml { .. }
            ) || matches!(
                &op,
                DomOp::SetAttribute(_, name, _) | DomOp::RemoveAttribute(_, name)
                    if matches!(name.as_str(), "src" | "srcset" | "poster" | "style")
            ) || matches!(
                &op,
                DomOp::SetStyleProperty(_, name, _)
                    if matches!(name.as_str(), "background" | "background-image" | "mask" | "mask-image")
            ) {
                entry.resources_dirty = true;
            }
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
                DomOp::CreateComment(data) => DomValue::Node(doc.create_comment(&data) as NodeId),

                // ── XML ──
                DomOp::CreateElementNS {
                    namespace,
                    qualified_name,
                    input_type,
                } => {
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
                DomOp::SetAttributeNS {
                    node,
                    namespace,
                    qualified_name,
                    value,
                } => {
                    doc.set_attribute_ns(to_hb(doc, node), &namespace, &qualified_name, &value);
                    DomValue::None
                }
                DomOp::GetAttributeNS {
                    node,
                    namespace,
                    local_name,
                } => match doc.get_attribute_ns(to_hb(doc, node), &namespace, &local_name) {
                    Some(v) => DomValue::Text(v),
                    None => DomValue::Null,
                },

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
                // webcore unlinks a child from whatever parent it is ACTUALLY
                // under, so there is no parent to redirect here — which is what
                // keeps the redirect on the way in from leaving a node linked
                // into the body and unlinked from the document.
                DomOp::RemoveChild { child, .. } => {
                    doc.remove_child(child as u32);
                    DomValue::Bool(true)
                }
                DomOp::InsertBefore {
                    parent,
                    child,
                    reference,
                } => {
                    let p = insertion_parent(doc, parent, child);
                    doc.insert_before(p, child as u32, reference as u32);
                    DomValue::Bool(true)
                }
                DomOp::ReplaceChild {
                    parent,
                    new_child,
                    old_child,
                } => {
                    let p = insertion_parent(doc, parent, new_child);
                    DomValue::Bool(doc.replace_child(p, new_child as u32, old_child as u32))
                }
                DomOp::CloneNode { node, deep } => {
                    let clone = doc.clone_node(to_hb(doc, node), deep);
                    if clone == 0 {
                        DomValue::Null
                    } else {
                        DomValue::Node(clone as NodeId)
                    }
                }
                // The document is not its own element. webcore has no node for
                // it, so `to_hb` answers `<html>` — right for reaching into the
                // tree, wrong for the two questions that ask what the node IS.
                // DOM §4.4: `9` and `#document`, which is what `widgets`
                // answers, and the seam is one API or it is not one.
                DomOp::NodeType(DOCUMENT) => DomValue::Number(9.0),
                DomOp::NodeName(DOCUMENT) => DomValue::Text("#document".to_string()),
                DomOp::NodeType(n) => DomValue::Number(f64::from(doc.node_type(to_hb(doc, n)))),
                DomOp::NodeName(n) => DomValue::Text(doc.node_name(to_hb(doc, n))),
                DomOp::NodeValue(n) => match doc.node_value(to_hb(doc, n)) {
                    Some(v) => DomValue::Text(v),
                    None => DomValue::Null,
                },
                DomOp::ParentNode(n) => {
                    let parent = doc.parent_node(to_hb(doc, n));
                    if parent == 0 {
                        DomValue::Null
                    } else {
                        DomValue::Node(from_hb(doc, parent))
                    }
                }
                DomOp::ChildNodes(n) => {
                    let kids = doc.child_nodes(to_hb(doc, n));
                    DomValue::Nodes(kids.iter().map(|c| from_hb(doc, *c)).collect())
                }
                DomOp::IsConnected(n) => DomValue::Bool(doc.is_connected(to_hb(doc, n))),
                DomOp::InnerHtml(n) => DomValue::Text(doc.inner_html(to_hb(doc, n))),
                // Parsing re-enters `apply` to build the tree, which would be a
                // second borrow of the document held here. Dispatched before
                // the lock instead.
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
                // does nothing; `widgets` answers the title, and one seam
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
                DomOp::GetAttribute(n, name) => match doc.get_attribute(to_hb(doc, n), &name) {
                    Some(v) => DomValue::Text(v),
                    None => DomValue::Null,
                },
                DomOp::AttributeNames(n) => DomValue::Texts(doc.get_attribute_names(to_hb(doc, n))),
                DomOp::RemoveAttribute(n, name) => {
                    doc.remove_attribute(to_hb(doc, n), &name);
                    DomValue::None
                }
                DomOp::SetStyleProperty(n, p, v) => {
                    doc.set_style_property(content_node(doc, n), &p, &v);
                    DomValue::None
                }
                // The DECLARED value — what was authored, un-resolved. webcore
                // already answered this way, which is why it disagreed with the
                // old `widgets` for `left`/`top`/`width`/`height`.
                DomOp::GetStyleProperty(n, p) => DomValue::Text(
                    doc.get_style_property(content_node(doc, n), &p)
                        .unwrap_or_default(),
                ),
                DomOp::StyleDeclarations(n) => {
                    let node = content_node(doc, n);
                    DomValue::Properties((0..doc.style_property_len(node))
                        .filter_map(|index| {
                            let name = doc.style_property_item(node, index)?;
                            let value = doc.get_style_property(node, &name)?;
                            Some((name, value))
                        }).collect())
                }
                // The RESOLVED value. Geometry comes off the laid-out rect;
                // everything else falls back to the declared value, matching
                // the floor `widgets` sets.
                DomOp::ComputedStyleProperty(n, p) => {
                    let node = content_node(doc, n);
                    DomValue::Text(doc.computed_style_property(node, &p))
                }

                // ── Form controls ──
                //
                // `checked` and `value` are CONTENT ATTRIBUTES in webcore, which
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
                DomOp::ObserveEvent { node, kind } => {
                    let key = (node, kind.clone());
                    if !entry.observed.contains_key(&key) {
                        let target = to_hb(doc, node);
                        let root = doc.root.node_id;
                        let sink = Arc::clone(&entry.targeted_events);
                        let listener = doc.add_event_listener(
                            target,
                            &kind,
                            Box::new(move |event, _| {
                                if let Ok(mut queue) = sink.lock() {
                                    let actual = if event.target == root {
                                        DOCUMENT
                                    } else {
                                        event.target as NodeId
                                    };
                                    queue.push_back(DomEventRecord {
                                        current_target: node,
                                        target: actual,
                                        kind: event.event_type.clone(),
                                        fields: UiEventFields {
                                            kind: event.event_type.clone(),
                                            key: event.key().to_owned(),
                                            code: event.code().to_owned(),
                                            client_x: event.client_x().round() as i32,
                                            client_y: event.client_y().round() as i32,
                                            button: event.button() as i32,
                                            buttons: event.buttons() as i32,
                                            delta_y: event.delta_y() as f64,
                                            ctrl_key: event.ctrl_key(),
                                            shift_key: event.shift_key(),
                                            alt_key: event.alt_key(),
                                            meta_key: event.meta_key(),
                                            repeat: event.repeat(),
                                            ..UiEventFields::default()
                                        },
                                    });
                                }
                            }),
                            webcore::dom::events::ListenerOptions::default(),
                        );
                        entry.observed.insert(key, listener);
                    }
                    DomValue::None
                }
                DomOp::UnobserveEvent { node, kind } => {
                    if let Some(listener) = entry.observed.remove(&(node, kind)) {
                        doc.remove_event_listener(listener);
                    }
                    DomValue::None
                }
                // Webcore hit-tests and dispatches to listeners attached at
                // their registered nodes. Form callbacks also report raw
                // control interaction for consumers without DOM listeners.
                DomOp::DispatchPointer {
                    kind,
                    client_x,
                    client_y,
                    button,
                } => {
                    use webcore::dom::HtmlEventType;
                    let etype = match kind.as_str() {
                        "mousedown" => HtmlEventType::MouseDown,
                        "mouseup" => HtmlEventType::MouseUp,
                        _ => HtmlEventType::MouseMove,
                    };
                    // `MouseEvent.button` is signed and `process_mouse_event`
                    // takes the same three values unsigned; anything else is
                    // not a button webcore knows and is treated as the primary
                    // one, which is what it does with an unrecognised device.
                    let button = u8::try_from(button).unwrap_or(0);
                    let point = (client_x + doc.scroll_x, client_y + doc.scroll_y);
                    let mut changed = if etype == HtmlEventType::MouseMove {
                        doc.dispatch_over_out(point)
                    } else {
                        false
                    };
                    changed |= doc.process_mouse_event(etype, point, button);
                    let pointer = match etype {
                        HtmlEventType::MouseDown => HtmlEventType::PointerDown,
                        HtmlEventType::MouseUp => HtmlEventType::PointerUp,
                        _ => HtmlEventType::PointerMove,
                    };
                    changed |= doc.process_mouse_event(pointer, point, button);
                    if etype == HtmlEventType::MouseUp && button == 2 {
                        changed |= doc.process_mouse_event(HtmlEventType::ContextMenu, point, button);
                    }
                    DomValue::Bool(changed)
                }
                DomOp::DispatchKeyboard(event) => {
                    use webcore::dom::HtmlEventType;
                    let kind = match event.kind.as_str() {
                        "keyup" => HtmlEventType::KeyUp,
                        "keypress" => HtmlEventType::KeyPress,
                        _ => HtmlEventType::KeyDown,
                    };
                    doc.process_key_event(
                        kind,
                        event.key_code as u32,
                        event.key.chars().next(),
                        event.ctrl_key,
                        event.shift_key,
                        event.alt_key,
                        event.meta_key,
                    );
                    DomValue::None
                }
                DomOp::DispatchWheel(event) => {
                    let client = (event.client_x as f32, event.client_y as f32);
                    let point = (client.0 + doc.scroll_x, client.1 + doc.scroll_y);
                    let mut wheel = webcore::dom::HtmlEvent::new(webcore::dom::HtmlEventType::Wheel);
                    wheel.client_pos = client;
                    wheel.doc_pos = point;
                    wheel.delta_y = event.delta_y as f32;
                    wheel.target = doc.element_from_point(client.0, client.1).unwrap_or(0);
                    let (_, wheel) = doc.dispatch_input_event(wheel);
                    let changed = !wheel.default_prevented
                        && doc.process_wheel_event_xy(point, 0.0, -(event.delta_y as f32));
                    DomValue::Bool(changed)
                }
                DomOp::DrainEvents => {
                    let mut targeted: Vec<DomEventRecord> = match entry.targeted_events.lock() {
                        Ok(mut q) => q.drain(..).collect(),
                        Err(_) => Vec::new(),
                    };
                    let fallback: Vec<(NodeId, String)> = match entry.events.lock() {
                        Ok(mut q) => q.drain(..).collect(),
                        Err(_) => Vec::new(),
                    };
                    for (node, kind) in fallback {
                        if !targeted.iter().any(|event| {
                            event.current_target == node && event.target == node && event.kind == kind
                        }) {
                            targeted.push(DomEventRecord {
                                current_target: node,
                                target: node,
                                fields: UiEventFields { kind: kind.clone(), ..UiEventFields::default() },
                                kind,
                            });
                        }
                    }
                    DomValue::Events(targeted)
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
                DomOp::DialogOpen(node) => DomValue::Bool(doc.dialog_open(to_hb(doc, node))),

                DomOp::BoundingClientRect(node) => {
                    // A geometry question flushes layout first — the whole
                    // reason `getBoundingClientRect` is specified to return a
                    // box rather than a cached number.
                    doc.flush_layout();
                    match doc.get_bounding_client_rect(to_hb(doc, node)) {
                        Some(rect) => DomValue::Rect {
                            x: rect.x as f64,
                            y: rect.y as f64,
                            width: rect.w as f64,
                            height: rect.h as f64,
                        },
                        None => DomValue::Rect {
                            x: 0.0,
                            y: 0.0,
                            width: 0.0,
                            height: 0.0,
                        },
                    }
                }
                DomOp::CanvasSize(node) => {
                    let node = to_hb(doc, node);
                    if !doc.local_name(node).eq_ignore_ascii_case("canvas") {
                        return DomValue::None;
                    }
                    let dimension = |name: &str, default: u32| doc.get_attribute(node, name)
                        .and_then(|value| value.trim().parse::<u32>().ok())
                        .unwrap_or(default);
                    DomValue::Pair(dimension("width", 300) as f64, dimension("height", 150) as f64)
                },

                // ── Not yet covered by this engine ──
                //
                // Each of these is a GAP, listed rather than quietly answered.
                // Nothing else falls here: of `DomOp`'s 64 variants these are
                // the only ones without an arm above, so a symptom that looks
                // like missing wiring is a bug in the arm, not an absence.
                //
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
            // webcore has no separate window object, and a `Document` IS the
            // browsing context a tab renders (which is what `browser.rs` does
            // with one `Document` per tab).
            WindowOp::Document(w) => WindowValue::Document(w),
            WindowOp::DefaultView(d) => WindowValue::Window(d),
            WindowOp::AdoptTopLevel(d) => WindowValue::Window(d),
            WindowOp::Closed(w) => WindowValue::Bool(
                docs()
                    .lock()
                    .map(|m| !m.entries.contains_key(&w))
                    .unwrap_or(true),
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
                let (width, height) = with_document(w, |doc| (doc.viewport_w, doc.viewport_h))
                    .unwrap_or((DEFAULT_VIEWPORT_W, DEFAULT_VIEWPORT_H));
                WindowValue::Pair(width as f64, height as f64)
            }
            WindowOp::Focus(_) => {
                webcore::embedded_window::focus();
                WindowValue::None
            }
            WindowOp::Screen(w) => {
                let size = webcore::embedded_window::screen_size().or_else(|| {
                    if let WindowValue::Pair(width, height) = self.window(WindowOp::InnerSize(w)) {
                        Some((width, height))
                    } else { None }
                });
                size.map(|(width, height)| WindowValue::Pair(width, height))
                    .unwrap_or(WindowValue::Null)
            }
            WindowOp::ResizeTo(w, width, height) => {
                with_document(w, |doc| doc.set_viewport(width as f32, height as f32));
                webcore::embedded_window::resize_to(width, height);
                WindowValue::None
            }
            WindowOp::ViewportChanged(w, width, height) => {
                with_document(w, |doc| doc.set_viewport(width as f32, height as f32));
                WindowValue::None
            }
            WindowOp::MoveTo(_, x, y) => {
                webcore::embedded_window::move_to(x, y);
                WindowValue::None
            }
            WindowOp::ScreenPosition(_) => webcore::embedded_window::screen_position()
                .map(|(x, y)| WindowValue::Pair(x, y))
                .unwrap_or(WindowValue::Null),
            WindowOp::Name(w) => {
                WindowValue::Text(with_document(w, |doc| doc.title()).unwrap_or_default())
            }
            WindowOp::Alert(message) => {
                webcore::platform::dialogs::alert(&message);
                WindowValue::None
            }
            WindowOp::Confirm(message) => {
                WindowValue::Bool(webcore::platform::dialogs::confirm(&message))
            }
        }
    }

    fn events(&self, op: EventOp) -> EventValue {
        match op {
            EventOp::Dispatch(event) => {
                webcore::ui_events::push(into_webcore_event(event));
                EventValue::None
            }
            EventOp::Poll => webcore::ui_events::poll()
                .map(from_webcore_event).map(EventValue::Event).unwrap_or(EventValue::Null),
            EventOp::Pending => EventValue::Count(webcore::ui_events::pending()),
            EventOp::PointerState => {
                let state = webcore::ui_events::pointer_state();
                EventValue::Pointer {
                    client_x: state.client_x, client_y: state.client_y,
                    buttons: state.buttons, ctrl_key: state.ctrl_key,
                    shift_key: state.shift_key, alt_key: state.alt_key,
                    meta_key: state.meta_key,
                }
            }
        }
    }

    fn schedule(&self, op: ScheduleOp) -> ScheduleValue {
        use webcore::scheduling;
        match op {
            ScheduleOp::SetTimer(delay) => ScheduleValue::Id(scheduling::set_timer(delay)),
            ScheduleOp::ClearTimer(id) => ScheduleValue::Bool(scheduling::clear_timer(id)),
            ScheduleOp::TakeDueTimer => scheduling::take_due_timer()
                .map(ScheduleValue::Id).unwrap_or(ScheduleValue::Null),
            ScheduleOp::RequestFrame => ScheduleValue::Id(scheduling::request_frame()),
            ScheduleOp::CancelFrame(id) => ScheduleValue::Bool(scheduling::cancel_frame(id)),
            ScheduleOp::TakeDueFrame => scheduling::take_due_frame()
                .map(ScheduleValue::Id).unwrap_or(ScheduleValue::Null),
            ScheduleOp::TimerDelayMs => scheduling::timer_delay_ms()
                .map(ScheduleValue::Ms).unwrap_or(ScheduleValue::Null),
            ScheduleOp::FrameDelayMs => scheduling::frame_delay_ms()
                .map(ScheduleValue::Ms).unwrap_or(ScheduleValue::Null),
            ScheduleOp::Now => ScheduleValue::Ms(scheduling::now_ms()),
        }
    }

    fn picker(&self, op: PickerOp) -> Vec<String> {
        use webcore::platform::dialogs;
        let paths = match op {
            PickerOp::Open { title, filters, directory, multiple ,} =>
                dialogs::open_file(&title, &filters, &directory, multiple),
            PickerOp::Save { title, filters, directory, suggested ,} =>
                dialogs::save_file(&title, &filters, &directory, &suggested).into_iter().collect(),
            PickerOp::Directory { title, directory } =>
                dialogs::pick_directory(&title, &directory).into_iter().collect(),
        };
        paths.into_iter().map(|path| path.to_string_lossy().into_owned()).collect()
    }
}

/// Install webcore as the web engine.
pub fn install() {
    crate::engine::set_engine(Arc::new(WebCore));
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn only_image_mutations_restart_resource_fetches() {
        let browser = WebCore;
        let document = browser.new_document("Image update");
        let node = match browser.document(document, DomOp::CreateElement {
            tag: "canvas".into(),
            input_type: String::new(),
        },) {
            DomValue::Node(node) => node,
            other => panic!("expected canvas node, got {other:?}"),
        };
        browser.document(document, DomOp::AppendChild { parent: DOCUMENT, child: node ,},);
        with_document_resources(document, |_, dirty| *dirty = false);

        browser.document(document, DomOp::SetStyleProperty(node, "color".into(), "red".into()),);
        with_document_resources(document, |_, dirty| assert!(!*dirty));

        browser.document(document, DomOp::SetStyleProperty(node, "background-image".into(), "url(second.png)".into()),);
        with_document_resources(document, |_, dirty| assert!(*dirty));
    }

    fn event_paths(events: &[DomEventRecord]) -> Vec<(NodeId, NodeId, &str)> {
        events.iter().map(|e| (e.current_target, e.target, e.kind.as_str())).collect()
    }

    #[test]
    fn inner_size_tracks_viewport_not_body_style() {
        let browser = WebCore;
        let document = browser.new_document("Viewport test");
        browser.document(document, DomOp::SetStyleProperty(DOCUMENT, "width".into(), "320px".into()),);
        let size = browser.window(WindowOp::InnerSize(document));
        assert!(matches!(size, WindowValue::Pair(800.0, 600.0)));

        browser.window(WindowOp::ResizeTo(document, 640.0, 880.0));
        let size = browser.window(WindowOp::InnerSize(document));
        assert!(matches!(size, WindowValue::Pair(640.0, 880.0)));
    }

    #[test]
    fn pointer_click_reaches_dom_listener_queue() {
        let browser = WebCore;
        let document = browser.new_document("Click test");
        let DomValue::Node(button) = browser.document(document, DomOp::CreateElement {
            tag: "button".into(), input_type: String::new(),
        },) else { panic!("button not created") };
        browser.document(document, DomOp::AppendChild { parent: DOCUMENT, child: button ,},);
        browser.document(document, DomOp::SetStyleProperty(button, "width".into(), "100px".into()),);
        browser.document(document, DomOp::SetStyleProperty(button, "height".into(), "40px".into()),);
        let DomValue::Rect { x, y, width, height ,} =
            browser.document(document, DomOp::BoundingClientRect(button))
        else { panic!("button not laid out") };
        assert!(width > 0.0 && height > 0.0);
        let (client_x, client_y) = ((x + width / 2.0) as f32, (y + height / 2.0) as f32);
        for kind in ["mousedown", "mouseup"] {
            browser.document(document, DomOp::DispatchPointer {
                kind: kind.into(), client_x, client_y, button: 0,
            },);
        }
        let DomValue::Events(events) = browser.document(document, DomOp::DrainEvents)
        else { panic!("no DOM event queue") };
        assert_eq!(event_paths(&events), vec![(button, button, "click")]);
    }

    #[test]
    fn toolbar_button_click_reaches_its_listener() {
        let browser = WebCore;
        let document = browser.new_document("Toolbar click test");
        let DomValue::Node(menu) = browser.document(
            document,
            DomOp::CreateElement {
                tag: "menu".into(),
                input_type: String::new(),
            },
        ) else {
            panic!("menu not created")
        };
        let DomValue::Node(button) = browser.document(
            document,
            DomOp::CreateElement {
                tag: "button".into(),
                input_type: String::new(),
            },
        ) else {
            panic!("button not created")
        };
        browser.document(
            document,
            DomOp::AppendChild {
                parent: DOCUMENT,
                child: menu,
            },
        );
        browser.document(
            document,
            DomOp::AppendChild {
                parent: menu,
                child: button,
            },
        );
        browser.document(
            document,
            DomOp::SetStyleProperty(menu, "width".into(), "210px".into()),
        );
        browser.document(
            document,
            DomOp::SetStyleProperty(menu, "height".into(), "20px".into()),
        );
        browser.document(
            document,
            DomOp::SetStyleProperty(button, "min-width".into(), "23px".into()),
        );
        browser.document(
            document,
            DomOp::SetStyleProperty(button, "height".into(), "100%".into()),
        );
        browser.document(
            document,
            DomOp::ObserveEvent {
                node: button,
                kind: "click".into(),
            },
        );

        let DomValue::Rect {
            x,
            y,
            width,
            height,
        } = browser.document(document, DomOp::BoundingClientRect(button))
        else {
            panic!("toolbar button not laid out")
        };
        assert!(width > 0.0 && height > 0.0);
        for kind in ["mousedown", "mouseup"] {
            browser.document(
                document,
                DomOp::DispatchPointer {
                    kind: kind.into(),
                    client_x: (x + width / 2.0) as f32,
                    client_y: (y + height / 2.0) as f32,
                    button: 0,
                },
            );
        }
        let DomValue::Events(events) = browser.document(document, DomOp::DrainEvents) else {
            panic!("no DOM event queue")
        };
        assert!(event_paths(&events).contains(&(button, button, "click")));
    }

    #[test]
    fn mouse_coordinates_reach_dom_listener_queue() {
        let browser = WebCore;
        let document = browser.new_document("Mouse fields test");
        let DomValue::Node(node) = browser.document(document, DomOp::CreateElement {
            tag: "div".into(), input_type: String::new(),
        },) else { panic!("div not created") };
        browser.document(document, DomOp::AppendChild { parent: DOCUMENT, child: node ,},);
        browser.document(document, DomOp::SetStyleProperty(node, "width".into(), "100px".into()),);
        browser.document(document, DomOp::SetStyleProperty(node, "height".into(), "40px".into()),);
        browser.document(document, DomOp::ObserveEvent { node, kind: "mousedown".into() ,},);
        let DomValue::Rect { x, y, width, height ,} =
            browser.document(document, DomOp::BoundingClientRect(node))
        else { panic!("div not laid out") };
        let client_x = (x + width / 2.0) as f32;
        let client_y = (y + height / 2.0) as f32;
        browser.document(document, DomOp::DispatchPointer {
            kind: "mousedown".into(), client_x, client_y, button: 0,
        },);
        let DomValue::Events(events) = browser.document(document, DomOp::DrainEvents)
        else { panic!("no DOM event queue") };
        assert_eq!(event_paths(&events), vec![(node, node, "mousedown")]);
        assert_eq!(events[0].fields.client_x, client_x.round() as i32);
        assert_eq!(events[0].fields.client_y, client_y.round() as i32);
    }

    #[test]
    fn dynamic_nested_div_click_reaches_dom_listener_queue() {
        let browser = WebCore;
        let document = browser.new_document("Div click test");
        let DomValue::Node(parent) = browser.document(document, DomOp::CreateElement {
            tag: "div".into(), input_type: String::new(),
        },) else { panic!("parent not created") };
        let DomValue::Node(child) = browser.document(document, DomOp::CreateElement {
            tag: "div".into(), input_type: String::new(),
        },) else { panic!("child not created") };
        browser.document(document, DomOp::AppendChild { parent: DOCUMENT, child: parent ,},);
        browser.document(document, DomOp::AppendChild { parent, child });
        browser.document(document, DomOp::ObserveEvent { node: parent, kind: "click".into() ,},);
        browser.document(document, DomOp::ObserveEvent { node: child, kind: "click".into() ,},);
        browser.document(document, DomOp::SetStyleProperty(parent, "width".into(), "100px".into()),);
        browser.document(document, DomOp::SetStyleProperty(parent, "height".into(), "60px".into()),);
        browser.document(document, DomOp::SetStyleProperty(child, "width".into(), "80px".into()),);
        browser.document(document, DomOp::SetStyleProperty(child, "height".into(), "40px".into()),);
        let DomValue::Rect { x, y, width, height ,} =
            browser.document(document, DomOp::BoundingClientRect(child))
        else { panic!("child not laid out") };
        assert!(width > 0.0 && height > 0.0);
        for kind in ["mousedown", "mouseup"] {
            browser.document(document, DomOp::DispatchPointer {
                kind: kind.into(), client_x: (x + width / 2.0) as f32,
                client_y: (y + height / 2.0) as f32, button: 0,
            },);
        }
        let DomValue::Events(events) = browser.document(document, DomOp::DrainEvents)
        else { panic!("no DOM event queue") };
        assert_eq!(event_paths(&events), vec![
            (child, child, "click"),
            (parent, child, "click"),
        ]);
        browser.document(document, DomOp::UnobserveEvent { node: child, kind: "click".into() ,},);
        with_entry(document, |entry| {
            assert!(!entry.observed.contains_key(&(child, "click".into())));
            assert!(!entry.doc.event_targets.node_ids().any(|id| id == child as u32));
        });
        for kind in ["mousedown", "mouseup"] {
            browser.document(document, DomOp::DispatchPointer {
                kind: kind.into(), client_x: (x + width / 2.0) as f32,
                client_y: (y + height / 2.0) as f32, button: 0,
            },);
        }
        let DomValue::Events(events) = browser.document(document, DomOp::DrainEvents)
        else { panic!("no DOM event queue") };
        // The legacy form callback still records the raw hit node; it has no
        // guest listener after removal, while the parent's native listener remains.
        assert_eq!(event_paths(&events), vec![
            (parent, child, "click"),
            (child, child, "click"),
        ]);
    }

    #[test]
    fn positioned_grid_cell_receives_pointer_click() {
        let browser = WebCore;
        let document = browser.new_document("Grid click test");
        let DomValue::Node(grid) = browser.document(document, DomOp::CreateElement {
            tag: "div".into(), input_type: String::new(),
        },) else { panic!("grid not created") };
        let DomValue::Node(cell) = browser.document(document, DomOp::CreateElement {
            tag: "div".into(), input_type: String::new(),
        },) else { panic!("cell not created") };
        browser.document(document, DomOp::ObserveEvent { node: cell, kind: "click".into() ,},);
        browser.document(document, DomOp::AppendChild { parent: DOCUMENT, child: grid ,},);
        browser.document(document, DomOp::AppendChild { parent: grid, child: cell ,},);
        for (name, value) in [
            ("display", "grid"), ("grid-template-columns", "repeat(7, 1fr)"),
            ("grid-template-rows", "repeat(6, 1fr)"), ("position", "absolute"),
            ("left", "50px"), ("top", "60px"), ("width", "400px"),
            ("height", "340px"), ("background-color", "dodgerblue"),
        ] {
            browser.document(document, DomOp::SetStyleProperty(grid, name.into(), value.into()),);
        }
        for (name, value) in [
            ("position", "relative"), ("width", "100%"), ("height", "100%"),
            ("grid-column-start", "1"), ("grid-row-start", "1"),
            ("background-color", "white"),
        ] {
            browser.document(document, DomOp::SetStyleProperty(cell, name.into(), value.into()),);
        }
        let DomValue::Rect { x, y, width, height ,} =
            browser.document(document, DomOp::BoundingClientRect(cell))
        else { panic!("cell not laid out") };
        assert!(width > 0.0 && height > 0.0);
        for kind in ["mousedown", "mouseup"] {
            browser.document(document, DomOp::DispatchPointer {
                kind: kind.into(), client_x: (x + width / 2.0) as f32,
                client_y: (y + height / 2.0) as f32, button: 0,
            },);
        }
        let DomValue::Events(events) = browser.document(document, DomOp::DrainEvents)
        else { panic!("no DOM event queue") };
        assert_eq!(event_paths(&events), vec![(cell, cell, "click")]);
    }

    #[test]
    fn input_families_reach_registered_dom_targets() {
        let browser = WebCore;
        let document = browser.new_document("Input events");
        let DomValue::Node(button) = browser.document(document, DomOp::CreateElement {
            tag: "button".into(), input_type: String::new(),
        },) else { panic!("button not created") };
        browser.document(document, DomOp::AppendChild { parent: DOCUMENT, child: button ,},);
        browser.document(document, DomOp::SetStyleProperty(button, "width".into(), "100px".into()),);
        browser.document(document, DomOp::SetStyleProperty(button, "height".into(), "40px".into()),);
        for kind in ["mousedown", "mouseup", "pointerdown", "pointerup", "click", "wheel",
                     "mousemove", "pointermove", "mouseover", "mouseout", "mouseenter", "mouseleave",] {
            browser.document(document, DomOp::ObserveEvent { node: button, kind: kind.into() ,},);
        }
        for kind in ["keydown", "keypress", "keyup"] {
            browser.document(document, DomOp::ObserveEvent { node: DOCUMENT, kind: kind.into() ,},);
        }
        let DomValue::Rect { x, y, width, height ,} =
            browser.document(document, DomOp::BoundingClientRect(button))
        else { panic!("button not laid out") };
        let (client_x, client_y) = ((x + width / 2.0) as f32, (y + height / 2.0) as f32);
        for kind in ["mousedown", "mouseup"] {
            browser.document(document, DomOp::DispatchPointer {
                kind: kind.into(), client_x, client_y, button: 0,
            },);
        }
        browser.document(document, DomOp::DispatchPointer {
            kind: "mousemove".into(), client_x, client_y, button: 0,
        },);
        let hover_before = with_document(document, |doc| doc.hovered_box).unwrap();
        let outside_hit = with_document(document, |doc| doc.element_from_point(700.0, 500.0)).unwrap();
        assert_eq!(hover_before, button as u32, "unexpected hover target; outside hit: {outside_hit:?}");
        browser.document(document, DomOp::DispatchPointer {
            kind: "mousemove".into(), client_x: 700.0,
            client_y: 500.0, button: 0,
        },);
        browser.document(document, DomOp::DispatchWheel(UiEventFields {
            kind: "wheel".into(), client_x: client_x as i32,
            client_y: client_y as i32, delta_y: -30.0,
            ..UiEventFields::default()
        }),);
        for kind in ["keydown", "keypress", "keyup"] {
            browser.document(document, DomOp::DispatchKeyboard(UiEventFields {
                kind: kind.into(), key: "a".into(), key_code: 65,
                ..UiEventFields::default()
            }),);
        }
        let DomValue::Events(events) = browser.document(document, DomOp::DrainEvents)
        else { panic!("no DOM event queue") };
        for kind in ["mousedown", "mouseup", "pointerdown", "pointerup", "click", "wheel",
                     "mousemove", "pointermove", "mouseover", "mouseout", "mouseenter", "mouseleave",] {
            assert!(events.iter().any(|event| event.current_target == button && event.kind == kind),
                "missing {kind}: {events:?}");
        }
        for kind in ["keydown", "keypress", "keyup"] {
            assert!(events.iter().any(|event| event.current_target == DOCUMENT && event.kind == kind),
                "missing {kind}: {events:?}");
        }
    }

    #[test]
    fn form_controls_emit_input_and_change() {
        let browser = WebCore;
        let document = browser.new_document("Form events");
        let DomValue::Node(input) = browser.document(document, DomOp::CreateElement {
            tag: "input".into(), input_type: "text".into(),
        },) else { panic!("input not created") };
        browser.document(document, DomOp::AppendChild { parent: DOCUMENT, child: input ,},);
        for kind in ["input", "change"] {
            browser.document(document, DomOp::ObserveEvent { node: input, kind: kind.into() ,},);
        }
        browser.document(document, DomOp::Focus(input));
        browser.document(document, DomOp::DispatchKeyboard(UiEventFields {
            kind: "keydown".into(), key: "a".into(), key_code: 65,
            ..UiEventFields::default()
        }),);
        let DomValue::Node(checkbox) = browser.document(document, DomOp::CreateElement {
            tag: "input".into(), input_type: "checkbox".into(),
        },) else { panic!("checkbox not created") };
        browser.document(document, DomOp::AppendChild { parent: DOCUMENT, child: checkbox ,},);
        browser.document(document, DomOp::ObserveEvent { node: checkbox, kind: "change".into() ,},);
        let DomValue::Rect { x, y, width, height ,} =
            browser.document(document, DomOp::BoundingClientRect(checkbox))
        else { panic!("checkbox not laid out") };
        for kind in ["mousedown", "mouseup"] {
            browser.document(document, DomOp::DispatchPointer {
                kind: kind.into(), client_x: (x + width / 2.0) as f32,
                client_y: (y + height / 2.0) as f32, button: 0,
            },);
        }
        let DomValue::Events(events) = browser.document(document, DomOp::DrainEvents)
        else { panic!("no DOM event queue") };
        assert!(events.iter().any(|event| event.current_target == input && event.kind == "input"),
            "input event missing: {events:?}");
        assert!(events.iter().any(|event| event.current_target == checkbox && event.kind == "change"),
            "change event missing: {events:?}");
    }

    #[test]
    fn populated_grid_hits_each_cell_not_just_its_container() {
        let browser = WebCore;
        let document = browser.new_document("Populated grid");
        let DomValue::Node(grid) = browser.document(document, DomOp::CreateElement {
            tag: "div".into(), input_type: String::new(),
        },) else { panic!("grid not created") };
        browser.document(document, DomOp::AppendChild { parent: DOCUMENT, child: grid ,},);
        for (name, value) in [
            ("display", "grid"), ("grid-template-columns", "repeat(7, 1fr)"),
            ("grid-template-rows", "repeat(6, 1fr)"), ("gap", "2px"),
            ("position", "absolute"), ("left", "50px"), ("top", "60px"),
            ("width", "400px"), ("height", "340px"),
        ] {
            browser.document(document, DomOp::SetStyleProperty(grid, name.into(), value.into()),);
        }
        browser.document(document, DomOp::SetAttribute(grid, "tabindex".into(), "1".into()),);
        for (tag, x, y, width, height) in [
            ("label", 0, 0, 500, 40),
            ("label", 50, 410, 400, 30),
            ("button", 202, 450, 96, 23),
        ] {
            let DomValue::Node(sibling) = browser.document(document, DomOp::CreateElement {
                tag: tag.into(), input_type: String::new(),
            },) else { panic!("sibling not created") };
            browser.document(document, DomOp::AppendChild { parent: DOCUMENT, child: sibling ,},);
            for (name, value) in [
                ("position", "absolute".to_string()),
                ("left", format!("{x}px")), ("top", format!("{y}px")),
                ("width", format!("{width}px")), ("height", format!("{height}px")),
            ] {
                browser.document(document, DomOp::SetStyleProperty(sibling, name.into(), value),);
            }
        }
        let mut cells = Vec::new();
        for row in 0..6 {
            for col in 0..7 {
                let DomValue::Node(cell) = browser.document(document, DomOp::CreateElement {
                    tag: "div".into(), input_type: String::new(),
                },) else { panic!("cell not created") };
                browser.document(document, DomOp::AppendChild { parent: grid, child: cell ,},);
                browser.document(document, DomOp::ObserveEvent { node: cell, kind: "click".into() ,},);
                for (name, value) in [
                    ("grid-column-start", (col + 1).to_string()),
                    ("grid-row-start", (row + 1).to_string()),
                    ("background-color", "white".into()),
                    ("border", "1px solid #777".into()),
                    ("margin", "2px 2px 2px 2px".into()),
                    ("position", "relative".into()),
                    ("width", "100%".into()),
                    ("height", "100%".into()),
                    ("box-sizing", "border-box".into()),
                ] {
                    browser.document(document, DomOp::SetStyleProperty(cell, name.into(), value));
                }
                cells.push(cell);
            }
        }
        for cell in cells {
            let DomValue::Rect { x, y, width, height ,} =
                browser.document(document, DomOp::BoundingClientRect(cell))
            else { panic!("cell not laid out") };
            assert!(width > 0.0 && height > 0.0, "cell {cell} has no hit area");
            let (client_x, client_y) = ((x + width / 2.0) as f32, (y + height / 2.0) as f32);
            browser.document(document, DomOp::DispatchPointer {
                kind: "mousemove".into(), client_x, client_y, button: 0,
            },);
            for kind in ["mousedown", "mouseup"] {
                browser.document(document, DomOp::DispatchPointer {
                    kind: kind.into(), client_x, client_y, button: 0,
                },);
            }
            let DomValue::Events(events) = browser.document(document, DomOp::DrainEvents)
            else { panic!("no DOM event queue") };
            assert!(events.iter().any(|event| event.current_target == cell && event.kind == "click"),
                "cell {cell} did not receive its click: {events:?}");
        }
    }
}
