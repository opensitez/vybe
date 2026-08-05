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
    CreateElement { tag: String, input_type: String },
    /// `document.getElementById(elementId)` — matches the `id` ATTRIBUTE.
    GetElementById(String),
    /// `document.querySelectorAll(tag)` — tag selectors.
    ElementsByTag(String),
    /// `document.title`
    Title,
    SetTitle(String),

    // ── Node ─────────────────────────────────────────────────────────────
    /// `parent.appendChild(child)`. A node has ONE parent, so this moves it.
    AppendChild { parent: NodeId, child: NodeId },
    RemoveChild { parent: NodeId, child: NodeId },
    /// `node.isConnected`
    IsConnected(NodeId),
    /// `node.textContent`
    TextContent(NodeId),
    SetTextContent(NodeId, String),

    // ── Element ──────────────────────────────────────────────────────────
    SetAttribute(NodeId, String, String),
    /// Absent yields [`DomValue::Null`], per spec — not an empty string.
    GetAttribute(NodeId, String),
    RemoveAttribute(NodeId, String),
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
    /// `select.add(option)` / `select.remove(index)` / clear.
    AddItem(NodeId, String),
    RemoveItem(NodeId, usize),
    ClearItems(NodeId),

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
    /// `(node, event type)` pairs from [`DomOp::DrainEvents`].
    Events(Vec<(NodeId, String)>),
}

// ── Windows: WHATWG HTML §7, browsing contexts ──────────────────────────

/// One `Window` operation.
#[derive(Clone, Debug)]
pub enum WindowOp {
    /// `window.open(url, target, features)` — creates the context AND its
    /// initial `about:blank` document.
    Open { target: String, features: String },
    /// `window.document`
    Document(WindowId),
    Close(WindowId),
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
    match engine() {
        Some(e) => e.document(document, op),
        None => DomValue::None }
}

pub fn window(op: WindowOp) -> WindowValue {
    match engine() {
        Some(e) => e.window(op),
        None => WindowValue::None }
}

pub fn events(op: EventOp) -> EventValue {
    match engine() {
        Some(e) => e.events(op),
        None => EventValue::None }
}

pub fn schedule(op: ScheduleOp) -> ScheduleValue {
    match engine() {
        Some(e) => e.schedule(op),
        None => ScheduleValue::None }
}

pub fn new_document(title: &str) -> DocumentId {
    match engine() {
        Some(e) => e.new_document(title),
        None => 0 }
}
