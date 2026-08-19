//! The seam between the `web:*` host functions and whatever implements them.
//!
//! Same shape, and the same reason, as [`canvas_backend`](crate::canvas_backend):
//! `Window`, `Document`, `Element` and `UIEvent` are web-platform interfaces,
//! so they are declared here; the machinery behind them belongs to an engine.
//! `vybe_widgets` is that engine today — its widget tree IS a document, with
//! nesting, per-node properties and events already in it — and a real browser
//! could be tomorrow, at which point the same guest code runs against the
//! browser's own DOM. Neither engine is named by this module.
//!
//! ONE trait for the whole engine rather than a seam per API family: windows,
//! documents and input are one implementation's job (`window.open` →
//! `document` → elements → events is a single chain), so splitting them would
//! only invite a half-swapped host. Canvas keeps its own seam because it is a
//! genuinely separate concern — a stream of paint ops, no queries.
//!
//! Enums rather than wide method sets so a backend implements THREE functions
//! and can never silently miss an operation: a missing arm is a compile error.
//! Unlike canvas ops these answer back — `createElement` yields a node,
//! `getAttribute` a string or null — so each `apply` returns a value.

use std::sync::{Arc, OnceLock, RwLock};

/// A document handle. One per browsing context — `window.document`.
pub type DocumentId = u64;

/// A browsing context handle — what `window.open` returns.
pub type WindowId = u64;

/// A node handle. `0` is the document element, i.e. `document.body`.
pub type NodeId = u64;

pub const DOCUMENT: NodeId = 0;

/// One DOM operation, in WHATWG terms.
#[derive(Clone, Debug)]
pub enum DomOp {
    // ── Document ─────────────────────────────────────────────────────────
    /// `document.createElement(localName)`, plus the `type` that — with the
    /// tag — is what HTML says decides which control an `<input>` is. The
    /// node is NOT inserted; it renders nothing until `appendChild`.
    CreateElement {
        tag: String,
        input_type: String,
    },
    /// `document.createTextNode(data)` — a `Text` node, uninserted.
    ///
    /// The DOM's other factory, and the half the seam never had: an element
    /// could be created and appended, but the CONTENT between two elements had
    /// no operation at all. Anything building a tree from markup needs it on
    /// the first text run it meets.
    CreateTextNode(String),
    /// `document.createComment(data)` — a `Comment` node, uninserted.
    ///
    /// HTML folds two more productions into this one node: `<![CDATA[…]]>` and
    /// `<?…>` are both comments outside foreign content (HTML §13.2.5.42), so
    /// this is the whole of what a parser needs beyond elements and text.
    CreateComment(String),
    /// `document.createCDATASection(data)` — XML's `<![CDATA[ … ]]>`.
    ///
    /// A `Text` node that spells itself differently (DOM §4.10), so its data
    /// counts towards an ancestor's `textContent`. HTML has no such
    /// production — it parses the same characters as a comment — so this is
    /// reachable only from an XML document.
    CreateCDataSection(String),
    /// `document.createElementNS(namespace, qualifiedName)` — DOM §4.5.
    ///
    /// A namespaced element is an ordinary element that also knows which
    /// vocabulary it came from; `prefix` and `localName` are two views of the
    /// qualified name, not two more fields, so neither crosses here.
    CreateElementNS {
        namespace: String,
        qualified_name: String,
        input_type: String,
    },
    /// `Element.namespaceURI` / `prefix` / `localName`.
    NamespaceUri(NodeId),
    Prefix(NodeId),
    LocalName(NodeId),
    /// `document.createProcessingInstruction(target, data)` — `<?target data?>`.
    /// XML-only, for the same reason, and NOT character data: excluded from an
    /// ancestor's `textContent` the way a comment is.
    CreateProcessingInstruction {
        target: String,
        data: String,
    },
    /// `canvas.width` / `canvas.height` — the BITMAP's size, in CSS pixels
    /// (HTML §4.12.5), defaulting to 300x150 when the attributes are absent.
    ///
    /// The read side of the content attributes `setAttribute` already writes.
    /// Without it a guest can size a surface and never ask how big it is, so
    /// the one operation every painter needs — "cover the whole surface" —
    /// cannot be spelled: `fillRect(0, 0, canvas.width, canvas.height)` is how
    /// a page clears to a colour, and there is no other way in the API.
    ///
    /// Answers `(width, height)` together because they are one fact and asking
    /// twice invites a caller to read a stale half.
    CanvasSize(NodeId),
    /// `document.getElementById(elementId)` — matches the `id` ATTRIBUTE.
    GetElementById(String),
    /// `document.getElementsByTagName(name)` — a tag NAME, not a selector.
    ElementsByTag(String),
    /// `document.querySelector(selectors)` — the first match in TREE order, or
    /// [`DomValue::Null`] when nothing matches.
    QuerySelector(String),
    /// `document.querySelectorAll(selectors)`, in tree order.
    ///
    /// A real selector, not a tag: the engine below the seam owns the matching,
    /// which is the point of the seam. An invalid selector yields an empty list
    /// rather than every element — the spec throws `SyntaxError` and there is
    /// no exception channel here, so it fails in the safe direction.
    QuerySelectorAll(String),
    /// `document.title`
    Title,
    SetTitle(String),

    // ── Node ─────────────────────────────────────────────────────────────
    /// `parent.appendChild(child)`. A node has ONE parent, so this moves it.
    AppendChild {
        parent: NodeId,
        child: NodeId,
    },
    RemoveChild {
        parent: NodeId,
        child: NodeId,
    },
    /// `parent.insertBefore(child, reference)`. A `reference` that is not a
    /// child answers [`DomValue::Bool`] `false` — the spec's `NotFoundError`,
    /// which has nowhere else to go here and must not become an append.
    InsertBefore {
        parent: NodeId,
        child: NodeId,
        reference: NodeId,
    },
    /// `parent.replaceChild(new_child, old_child)` — the new node takes the
    /// old one's POSITION, which is the whole difference from removing and
    /// appending.
    ReplaceChild {
        parent: NodeId,
        new_child: NodeId,
        old_child: NodeId,
    },
    /// `node.cloneNode(deep)` — a copy that is NOT in the document, exactly as
    /// `createElement`'s result is not.
    CloneNode {
        node: NodeId,
        deep: bool,
    },
    /// `node.nodeType` / `nodeName` / `nodeValue` / `parentNode` /
    /// `childNodes` — the read side of a node, DOM §4.4.
    ///
    /// **Operations, not properties on a handle.** A handle is minted once and
    /// a tree changes: `childNodes` stamped at creation is right when it is
    /// taken and wrong immediately after the next `appendChild`, which is the
    /// duplicated state this seam exists to remove. Immutable facts may ride
    /// on a handle; these may not.
    NodeType(NodeId),
    NodeName(NodeId),
    NodeValue(NodeId),
    ParentNode(NodeId),
    ChildNodes(NodeId),
    /// `element.innerHTML` — read the subtree as markup.
    InnerHtml(NodeId),
    /// `element.innerHTML = …` — REPLACE the subtree from markup.
    SetInnerHtml {
        node: NodeId,
        html: String,
    },
    /// `node.isConnected`
    IsConnected(NodeId),
    /// `node.textContent`
    TextContent(NodeId),
    SetTextContent(NodeId, String),

    // ── Element ──────────────────────────────────────────────────────────
    SetAttribute(NodeId, String, String),
    /// Absent yields [`DomValue::Null`], per spec — not an empty string.
    GetAttribute(NodeId, String),
    /// `element.getAttributeNames()` — DOM §4.9. The qualified names of the
    /// element's content attributes.
    AttributeNames(NodeId),
    RemoveAttribute(NodeId, String),
    /// `element.setAttributeNS(namespace, qualifiedName, value)` and
    /// `getAttributeNS(namespace, localName)` — DOM §4.9.
    ///
    /// Note the asymmetry, which is the spec's and not a slip: the write takes
    /// a QUALIFIED name (that is what serialises) and the read takes a LOCAL
    /// one (matched together with the namespace). `xlink:href` and `href`
    /// share a local name and are two attributes.
    SetAttributeNS {
        node: NodeId,
        namespace: String,
        qualified_name: String,
        value: String,
    },
    GetAttributeNS {
        node: NodeId,
        namespace: String,
        local_name: String,
    },
    /// `element.style.setProperty(property, value)` — CSS text with units.
    SetStyleProperty(NodeId, String, String),
    GetStyleProperty(NodeId, String),
    /// `element.focus()`
    Focus(NodeId),

    // ── HTMLInputElement / HTMLSelectElement IDL ─────────────────────────
    /// `input.value` — a string, whatever the control underneath.
    Value(NodeId),
    SetValue(NodeId, String),
    /// `input.checked` — a boolean, not the string `"True"`.
    Checked(NodeId),
    SetChecked(NodeId, bool),
    /// `select.selectedIndex` — the index of the first selected option, or
    /// `-1` when nothing is selected. Its own IDL member, NOT `value`:
    /// `HTMLSelectElement.value` is a DOMString (the selected option's value),
    /// and a control that answered an index there would not survive contact
    /// with a browser.
    SelectedIndex(NodeId),
    SetSelectedIndex(NodeId, i32),
    /// `select.options[index].text` — the text of one option.
    /// Out of range answers `""`, not a trap.
    ItemText(NodeId, usize),
    SetItemText(NodeId, usize, String),
    /// `select.add(option)` / `select.remove(index)` / clear.
    AddItem(NodeId, String),
    RemoveItem(NodeId, usize),
    ClearItems(NodeId),

    // ── HTMLDialogElement ────────────────────────────────────────────────
    /// `dialog.show()` when `modal` is false, `dialog.showModal()` when true.
    ///
    /// Both set the `open` attribute — that is what "open" MEANS for a
    /// dialog, and it is a reflected attribute, so it belongs in the tree
    /// rather than beside it. `modal` is carried rather than flattened away
    /// because the UA stylesheet gives `dialog:modal` a different box:
    /// `position: fixed` against the viewport instead of in-flow.
    ///
    /// What `modal` does NOT get here is the top layer or the inertness of
    /// everything behind it. Both are renderer behaviour with nowhere to live
    /// in a tree — `:modal` has no attribute to carry them — so a modal
    /// dialog currently renders as an open, viewport-positioned one.
    ShowDialog {
        node: NodeId,
        modal: bool,
    },
    /// `dialog.close()` — clears `open`. The `close` EVENT and `returnValue`
    /// are the surface's job: neither is a reflected attribute.
    CloseDialog(NodeId),
    /// `dialog.open` — the reflected boolean attribute, read back.
    DialogOpen(NodeId),

    /// `input.showPicker()` — HTML's "show the picker, if applicable" for the
    /// input types that own one (`file`, `color`, `date`, …). What a picker
    /// IS belongs to the engine; there is no second colour dialog here.
    ShowPicker(NodeId),

    /// Drain what the user did, as DOM events (`click`, `input`, `change`).
    /// The surface turns each into an `Event` object and calls the listeners
    /// registered on that node.
    DrainEvents,
}

/// What an operation answers.
#[derive(Clone, Debug)]
pub enum DomValue {
    None,
    /// An absent attribute — `null`, distinct from `""`.
    Null,
    Node(NodeId),
    Nodes(Vec<NodeId>),
    Text(String),
    Bool(bool),
    /// A numeric IDL attribute — `select.selectedIndex`. Distinct from
    /// [`DomValue::Text`] because the IDL type is `long`, and a caller that
    /// compares it with `>= 0` needs a number rather than digits.
    Number(f64),
    /// `(node, event type)` pairs from [`DomOp::DrainEvents`].
    Events(Vec<(NodeId, String)>),
    /// Two numbers that are one fact — a size, today.
    Pair(f64, f64),
    /// A list of strings — attribute names, today.
    Texts(Vec<String>),
}

// ── Windows: WHATWG HTML §7, browsing contexts ──────────────────────────

/// One `Window` operation.
#[derive(Clone, Debug)]
pub enum WindowOp {
    /// `window.open(url, target, features)` — creates the context AND its
    /// initial `about:blank` document.
    Open {
        target: String,
        features: String,
    },
    /// `window.document`
    Document(WindowId),
    /// `Document.defaultView` — the inverse of [`WindowOp::Document`].
    ///
    /// A document reaches its own window through this in every real engine; a
    /// script that holds the global uses `window`/`self` instead, which a guest
    /// with no global object cannot spell.
    DefaultView(DocumentId),
    /// Give an existing document its TOP-LEVEL browsing context — the tab the
    /// user agent makes before any script runs.
    ///
    /// Not a spec operation, because in a browser it is not an operation at
    /// all: the context is already there. It exists because a standalone
    /// toolkit has to bootstrap the thing a user agent is handed.
    AdoptTopLevel(DocumentId),
    Close(WindowId),
    /// `window.focus()` — bring the context to the front.
    Focus(WindowId),
    /// `window.screen` (CSSOM View) — the DISPLAY's size, not the window's.
    Screen(WindowId),
    /// `window.closed`
    Closed(WindowId),
    /// `window.innerWidth` / `innerHeight`
    InnerSize(WindowId),
    ResizeTo(WindowId, f64, f64),
    MoveTo(WindowId, f64, f64),
    /// `window.screenX` / `screenY`
    ScreenPosition(WindowId),
    /// `window.name`
    Name(WindowId),
    /// `window.alert(message)` — a modal one-button box. Blocking is not an
    /// implementation shortcut here, it IS the spec: `alert` suspends script
    /// until dismissed, and so does the toolkit call every frontend spells it
    /// with (`ShowMessage`, `MessageBox.Show`).
    ///
    /// No `WindowId`: `alert` is invoked on the global object in every real
    /// page, and a guest that has never opened a window still has one to talk
    /// to. Adding an id would make the common call the awkward one.
    Alert(String),
    /// `window.confirm(message)` — two buttons, `true` for OK.
    Confirm(String),
}

#[derive(Clone, Debug)]
pub enum WindowValue {
    None,
    Null,
    Window(WindowId),
    Document(DocumentId),
    Bool(bool),
    Text(String),
    /// A `(width, height)` or `(x, y)` pair.
    Pair(f64, f64),
}

// ── Input: W3C UI Events ────────────────────────────────────────────────

/// One UI-event-queue operation. The event itself crosses this seam as its
/// spec fields rather than as a struct, so the engine and the host never have
/// to agree on a Rust type.
#[derive(Clone, Debug)]
pub enum EventOp {
    /// `EventTarget.dispatchEvent(event)` — inject a synthetic event, exactly
    /// as the DOM allows, which also makes the pipeline testable with no
    /// window open.
    Dispatch(UiEventFields),
    /// The drain a polling guest needs where a browser invokes listeners.
    Poll,
    /// Queue depth, so a loop can drain without polling blind.
    Pending,
    /// What a page tracks from the event stream, for guests that sample.
    PointerState,
}

/// A UI event's spec fields — W3C attribute names, no VM types.
#[derive(Clone, Debug, Default)]
pub struct UiEventFields {
    pub kind: String,
    pub key: String,
    pub code: String,
    pub key_code: i32,
    pub client_x: i32,
    pub client_y: i32,
    pub button: i32,
    pub buttons: i32,
    pub delta_y: f64,
    pub ctrl_key: bool,
    pub shift_key: bool,
    pub alt_key: bool,
    pub meta_key: bool,
    pub repeat: bool,
}

#[derive(Clone, Debug)]
pub enum EventValue {
    None,
    /// No event was queued — `pollEvent` answers null.
    Null,
    Count(usize),
    Event(UiEventFields),
    /// `{clientX, clientY, buttons, ctrlKey, shiftKey, altKey, metaKey}`.
    Pointer {
        client_x: i32,
        client_y: i32,
        buttons: i32,
        ctrl_key: bool,
        shift_key: bool,
        alt_key: bool,
        meta_key: bool,
    },
}

// ── Scheduling: HTML timers + the animation frame clock ─────────────────

/// One scheduling operation. These deal in IDS, never callbacks: the engine
/// says what became due, and the callback registry above decides what running
/// it means — the same division the DOM uses for listeners.
#[derive(Clone, Debug)]
pub enum ScheduleOp {
    /// `setTimeout` / `setInterval` — schedule `delay_ms` from now.
    SetTimer(f64),
    /// `clearTimeout` / `clearInterval`.
    ClearTimer(u64),
    /// Pop ONE due timer id, first-registered-due-first.
    TakeDueTimer,
    /// `requestAnimationFrame`.
    RequestFrame,
    /// `cancelAnimationFrame`.
    CancelFrame(u64),
    /// Pop ONE registration due this frame.
    TakeDueFrame,
    /// How long until the earliest timer / next frame, RELATIVE.
    ///
    /// Relative because the engine's monotonic clock has its own origin; an
    /// absolute timestamp from it would mis-schedule every sleep the host
    /// computes against `event_loop::monotonic_now_ms`.
    TimerDelayMs,
    FrameDelayMs,
    /// `performance.now()` — the timestamp a frame callback is handed.
    Now,
}

#[derive(Clone, Debug)]
pub enum ScheduleValue {
    None,
    /// Nothing due / nothing scheduled.
    Null,
    Id(u64),
    Bool(bool),
    /// A relative delay, or a timestamp, in milliseconds.
    Ms(f64),
}

/// The engine behind the `web:*` surface — windows, documents, input,
/// scheduling.
pub trait WebEngine: Send + Sync {
    /// Open a document directly, without a browsing context. A guest that
    /// never calls `window.open` still has one document to build into.
    fn new_document(&self, title: &str) -> DocumentId;

    /// Open an XML document. The one thing that genuinely differs from an
    /// HTML one down here is CASE: HTML folds tag and attribute names, XML
    /// keeps them, and that is a property of the document rather than of any
    /// call on it — which is the distinction a browser draws between
    /// `Document` and `XMLDocument`.
    fn new_xml_document(&self, title: &str) -> DocumentId;
    fn document(&self, document: DocumentId, op: DomOp) -> DomValue;
    fn window(&self, op: WindowOp) -> WindowValue;
    fn events(&self, op: EventOp) -> EventValue;
    fn schedule(&self, op: ScheduleOp) -> ScheduleValue;
}

fn slot() -> &'static RwLock<Option<Arc<dyn WebEngine>>> {
    static ENGINE: OnceLock<RwLock<Option<Arc<dyn WebEngine>>>> = OnceLock::new();
    ENGINE.get_or_init(|| RwLock::new(None))
}

/// Install the engine. Called once at startup by whoever supplies it.
pub fn set_engine(engine: Arc<dyn WebEngine>) {
    *slot().write().unwrap() = Some(engine);
}

pub fn engine() -> Option<Arc<dyn WebEngine>> {
    slot().read().unwrap().clone()
}

/// Apply a document operation, or `None` when no engine is installed — a
/// headless run has no document, exactly as a canvas op with no painter
/// draws nothing.
pub fn apply(document: DocumentId, op: DomOp) -> DomValue {
    // **`innerHTML =` is parsed BEFORE the engine is entered.** Building the
    // fragment means calling `apply` again, once per element, and the engine
    // holds the document for the length of one op — so dispatching this
    // through `e.document(..)` like every other write would deadlock on the
    // first tag. Handling it here keeps each of those inner writes an ordinary,
    // separately-locked operation.
    if let DomOp::SetInnerHtml { node, html } = &op {
        crate::dom_parser::set_inner_html(document, *node, html);
        return DomValue::None;
    }
    match engine() {
        Some(e) => e.document(document, op),
        None => DomValue::None,
    }
}

pub fn window(op: WindowOp) -> WindowValue {
    match engine() {
        Some(e) => e.window(op),
        None => WindowValue::None,
    }
}

pub fn events(op: EventOp) -> EventValue {
    match engine() {
        Some(e) => e.events(op),
        None => EventValue::None,
    }
}

pub fn schedule(op: ScheduleOp) -> ScheduleValue {
    match engine() {
        Some(e) => e.schedule(op),
        None => ScheduleValue::None,
    }
}

pub fn new_document(title: &str) -> DocumentId {
    match engine() {
        Some(e) => e.new_document(title),
        None => 0,
    }
}

/// Open an XML document — see [`WebEngine::new_xml_document`].
pub fn new_xml_document(title: &str) -> DocumentId {
    match engine() {
        Some(e) => e.new_xml_document(title),
        None => 0,
    }
}
