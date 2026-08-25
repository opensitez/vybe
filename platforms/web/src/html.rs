//! WHATWG DOM, HTML and CSSOM — split by which SPEC defines each operation,
//! not by convenience:
//!
//! - **`web:dom`** (DOM Standard) — `createElement`, `appendChild`,
//!   `removeChild`, `isConnected`, `textContent`, `get`/`set`/`removeAttribute`,
//!   `getElementById`, `getElementsByTagName`, `addEventListener`.
//! - **`web:html`** (HTML Standard) — the parts HTML adds to `Document` and
//!   `Element`: `body`, `title`, and the element IDL `value`, `checked`,
//!   `focus`, and `HTMLSelectElement`'s option list.
//! - **`web:cssom`** (CSSOM) — `element.style.setProperty` / `getPropertyValue`.
//!
//! A guest reaches all of it through `window.document`, which is where a
//! browser puts it too.
//!
//! This module is exposure, not implementation. The document lives in the
//! engine behind [`dom_backend`](crate::dom_backend); everything here does is
//! turn host calls into [`DomOp`]s and spec-shaped results back — so the same
//! surface works over the widget toolkit today and over a real browser's DOM
//! later, and so the toolkit stays usable on its own, with no runtime under it.
//!
//! What that leaves this file responsible for is exactly the web semantics a
//! relabelled toolkit API would lose:
//!
//! - `createElement` returns a node that is NOT in the document — it has no
//!   parent and renders nothing until `appendChild` inserts it.
//! - `getAttribute` on an absent attribute is `null`, never `""`.
//! - `checked` is a boolean and `value` a string, never `"True"`.
//! - a listener receives an `Event` OBJECT (`type`, `target`), not bare
//!   arguments — and `click`/`input`/`change` are the event names.
//! - every entry point is scoped to a document, because `createElement` is a
//!   method on `window.document` and a guest with two windows has two trees.
//!
//! Listeners are the one thing kept here rather than in the engine: a
//! registration holds a VM callback, which is a runtime value the toolkit has
//! no business knowing about.

use std::collections::HashMap;
use std::sync::{Arc, Mutex};

use vybe_runtime::value::Object;
use vybe_runtime::vm::{HostFnDecl, ResourceBinding, ResourceMemberKind};
use vybe_runtime::{FuncSig, HostContext, VM, ValType, Value};

use crate::engine::{DOCUMENT, DocumentId, DomOp, DomValue, NodeId, apply};

/// `(document, node, event type)` → listeners, in registration order.
type ListenerKey = (DocumentId, NodeId, String);

/// A named type, not a bare `HashMap`: [`vybe_runtime::resources`] keys by
/// `TypeId`, so two plugins storing the same std type would share one cell.
#[derive(Default)]
struct Listeners(HashMap<ListenerKey, Vec<Value>>);

impl std::ops::Deref for Listeners {
    type Target = HashMap<ListenerKey, Vec<Value>>;
    fn deref(&self) -> &Self::Target {
        &self.0
    }
}
impl std::ops::DerefMut for Listeners {
    fn deref_mut(&mut self) -> &mut Self::Target {
        &mut self.0
    }
}

/// Every DOM event listener the running program registered — guest `Value`s,
/// so a listener surviving a reset leaves one program's handlers wired for the
/// next. VM-owned ([`vybe_runtime::resources`]): dropped on `reset_to`.
fn listeners() -> &'static Mutex<Listeners> {
    vybe_runtime::resources::get::<Listeners>()
}

/// Every `(node, type, callback)` registered on a document, for a host that
/// needs to see what the guest wired up — a form launcher deciding whether a
/// window is interactive, or a test asserting a handler landed.
///
/// Read-only and ordered per key: `addEventListener` appends, and a listener
/// list is a sequence, not a set.
pub fn document_listeners(document: DocumentId) -> Vec<(NodeId, String, Value)> {
    listeners()
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
        .iter()
        .filter(|((doc, _, _), _)| *doc == document)
        .flat_map(|((_, node, kind), callbacks)| {
            callbacks
                .iter()
                .map(move |callback| (*node, kind.clone(), callback.clone()))
        })
        .collect()
}

/// `dialog.returnValue` per dialog.
///
/// Kept beside the tree rather than in it for the same reason listeners are:
/// `returnValue` is NOT a reflected attribute. HTML gives it no attribute to
/// reflect to, so a `<dialog>` that has closed with a value serialises exactly
/// like one that has not — putting it in the tree would invent markup the
/// spec does not have.
#[derive(Default)]
struct ReturnValues(HashMap<(DocumentId, NodeId), String>);

fn return_values() -> &'static Mutex<ReturnValues> {
    vybe_runtime::resources::get::<ReturnValues>()
}

/// Drop every listener registered on a document — the counterpart to
/// [`reset_active_document`], since a discarded document's listeners are
/// unreachable but the map still holds them.
pub fn clear_document_listeners(document: DocumentId) {
    listeners()
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
        .retain(|(doc, _, _), _| *doc != document);
}

/// `target.addEventListener(type, callback)`.
pub fn add_event_listener(document: DocumentId, node: NodeId, kind: &str, callback: Value) {
    // ⛔ **An event type is CASE-SENSITIVE** (DOM §2.7 — the type is a plain
    // string and listeners are keyed by it). `addEventListener("Click", …)`
    // never fires for a real click in a browser, and folding it here made the
    // two the same event — leniency that lets a frontend register the wrong
    // spelling and still appear to work, which is exactly what hid
    // `onChanged` being wired to `click`.
    //
    // The debugger's `fire <control> Click` used to depend on this fold; it now
    // folds at its own edge, where a convenience for a person typing a command
    // belongs, rather than in the DOM.
    listeners()
        .lock()
        .unwrap()
        .entry((document, node, kind.to_string()))
        .or_default()
        .push(callback);
}

/// `EventTarget.removeEventListener` — DOM §2.7.
///
/// **Identity, not equality.** The spec removes the listener whose callback is
/// *the same object*, and `Value`'s `==` cannot express that: two
/// `ObjectKind::Function`s compare equal when their `chunk_index` matches, so
/// every closure produced by one factory is "equal" to its siblings. A page
/// doing `makeHandler(d)` in a loop — the calculator does exactly this — would
/// have the wrong key unsubscribed, silently. `Arc::ptr_eq` is the identity the
/// spec actually means.
///
/// A callback that is not an object cannot be identified, so it matches
/// nothing and removes nothing. Refusing to guess is the safe half: removing
/// the wrong listener is invisible, and removing none is at worst a listener
/// that keeps firing.
pub fn remove_event_listener(document: DocumentId, node: NodeId, kind: &str, callback: &Value) {
    let mut all = listeners().lock().unwrap();
    let key = (document, node, kind.to_string());
    let Some(list) = all.get_mut(&key) else {
        return;
    };
    // Only the FIRST match goes, per spec: adding the same callback twice is a
    // no-op the second time, so there is never more than one to find — and if
    // an earlier path did double-register, removing one per call is still what
    // a browser does.
    // **Two tiers, because two languages mean two different things.**
    //
    // 1. OBJECT IDENTITY — what JS means. `removeEventListener` matches the
    //    reference the program kept, and a `bind` result is a different object
    //    (measured against node: `f.bind(null) !== f`, as ECMA requires). A
    //    browser cannot remove a bound listener either without that reference.
    // 2. DELEGATE EQUALITY — what `RemoveHandler`/`Align`-era frontends mean.
    //    .NET removes a handler by target+method and VCL's method pointer is
    //    the same `(Self, code)` pair; neither is object identity, and neither
    //    frontend ever sees the wrapper `emit_gui_property_set` bound on the
    //    way in. `Value::eq` compares two functions by `chunk_index`, which IS
    //    that equality — and the bound wrapper inherits its target's kind, so
    //    it matches the method the program named.
    //
    // Identity is tried across the WHOLE list first. Falling back per-element
    // would let tier 2 claim a sibling closure while the exact object sat
    // later in the list — the wrong listener removed, invisibly.
    let found = list
        .iter()
        .position(|held| same_callback(held, callback))
        .or_else(|| {
            list.iter()
                .position(|held| is_bound_wrapper(held) && held.eq(callback))
        });
    if let Some(i) = found {
        list.remove(i);
    }
    if list.is_empty() {
        all.remove(&key);
    }
}

/// Was this listener produced by `Function.prototype.bind`?
///
/// The scope of tier 2, and the reason it is safe. A bound wrapper is the ONLY
/// listener a frontend cannot name: `emit_gui_property_set` binds the receiver
/// in on the way to `addEventListener`, so the program's `AddressOf OnClick`
/// never was the stored object. Everything else the program CAN name, and for
/// those, identity is the whole answer.
///
/// Without this scope, a page removing a listener it never added would evict a
/// same-bodied sibling — measured, by the test that pins it.
fn is_bound_wrapper(value: &Value) -> bool {
    let Value::Object(obj) = value else {
        return false;
    };
    obj.lock()
        .map(|o| o.properties.contains_key("__bound_args"))
        .unwrap_or(false)
}

/// Are these two values the SAME callback object?
fn same_callback(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Object(x), Value::Object(y)) => std::sync::Arc::ptr_eq(x, y),
        _ => false,
    }
}

pub fn listeners_for(document: DocumentId, node: NodeId, kind: &str) -> Vec<Value> {
    listeners()
        .lock()
        .unwrap()
        .get(&(document, node, kind.to_string()))
        .cloned()
        .unwrap_or_default()
}

/// Drain user interaction into `(callback, event object)` pairs ready to
/// invoke. The frame loop calls this and dispatches each pair into the VM —
/// the callback is a runtime value, so invoking it is the runtime's job.
pub fn pending_dispatches(document: DocumentId) -> Vec<(Value, Value)> {
    let DomValue::Events(events) = apply(document, DomOp::DrainEvents) else {
        return Vec::new();
    };
    let mut out = Vec::new();
    for (node, kind) in events {
        for cb in listeners_for(document, node, &kind) {
            out.push((cb, event_object(&kind, node)));
        }
    }
    out
}

/// The object a listener receives — a real `Event`, not loose arguments.
pub fn event_object(kind: &str, target: NodeId) -> Value {
    let mut o = Object::new();
    o.properties
        .insert("type".into(), Value::String(kind.into()));
    o.properties
        .insert("target".into(), Value::F64(target as f64));
    o.properties
        .insert("currentTarget".into(), Value::F64(target as f64));
    o.properties.insert("bubbles".into(), Value::Bool(true));
    o.properties.insert("cancelable".into(), Value::Bool(true));
    Value::Object(vybe_runtime::heap::alloc(o))
}

/// The active document, created on first use — one per AGENT, not per
/// process.
///
/// HTML §7 puts a browsing context inside an agent, and an agent is one
/// thread of execution: two guests running side by side are two agents and
/// must not see each other's tree. A process-wide document made them share
/// one, which is invisible to a single-window program and wrong the moment
/// anything runs two (a test binary being the obvious case, where every test
/// would accumulate the previous one's controls).
///
/// `window.open` still creates additional documents and hands back their
/// handles explicitly — this is the ambient one, not the only one.
pub fn active_document() -> DocumentId {
    let slot = active_document_slot();
    let mut slot = slot.lock().unwrap();
    match slot.0 {
        Some(id) => id,
        None => {
            let id = crate::engine::new_document("");
            // **The tab, made before any script runs.** A document without a
            // browsing context is a thing a browser never has, and every window
            // verb was unreachable without one: `Document.defaultView` answered
            // null for every program, so nothing could name the window to close
            // it, move it, or ask its size. `adopt` is idempotent, and this is
            // the one place the AMBIENT document comes into being — `open()`
            // brings its own context with it.
            crate::engine::window(crate::engine::WindowOp::AdoptTopLevel(id));
            slot.0 = Some(id);
            id
        }
    }
}

/// This browsing context's ambient document.
///
/// VM-owned ([`vybe_runtime::resources`]) so `reset_to` drops it: navigating
/// away is what a browsing context does, and a slot that never resets is how a
/// test on a reused thread inherited the previous test's controls — a 2-test
/// wobble between identical runs before this was resettable at all.
#[derive(Default)]
struct ActiveDocument(Option<DocumentId>);

fn active_document_slot() -> &'static std::sync::Mutex<ActiveDocument> {
    vybe_runtime::resources::get::<ActiveDocument>()
}

/// Whether this agent has a BROWSING CONTEXT (HTML §7.1).
///
/// A browsing context is the window. A browser that opens one has a tab on
/// screen before a single byte of content arrives, and `about:blank` is a
/// window with nothing in it — so "is there a window" is answered by whether a
/// context was opened, never by what it contains.
///
/// Deliberately does NOT create one, unlike [`active_document`]: asking the
/// question must not answer it. That is why this reads the slot directly.
///
/// The slot is an honest proxy for the context now, and was not always: until
/// the ambient document was adopted into a top-level context, NOTHING created a
/// `BrowsingContext` except `window.open()`, so this function was named for one
/// concept and testing another. Setting the slot and adopting the context are
/// the same step in [`active_document`], so the two cannot disagree.
pub fn has_browsing_context() -> bool {
    active_document_slot()
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
        .0
        .is_some()
}

// ── argument decoding ───────────────────────────────────────────────────

fn doc_arg_id(args: &[Value], idx: usize) -> Option<DocumentId> {
    args.get(idx).map(|v| v.as_f64() as DocumentId)
}

/// The document a call is about.
///
/// Two spellings, because a document is reached two ways. `activeDocument`
/// answers a bare id and every ambient call passes that number straight back.
/// A document that is NOT the ambient one — `window.open`'s, and now
/// `DOMParser.parseFromString`'s — is held as a handle, and the id travels
/// inside it as `__document`, exactly as a node's travels as `__node`.
///
/// Without the second spelling a parsed document is unreachable: it exists,
/// it has a tree, and no operation can name it.
///
/// **Id `0` means THE ACTIVE DOCUMENT**, not "no document". Real ids start at 1
/// (`new_document` increments before handing one out), so 0 was never a
/// document and every call carrying it used to silently do nothing — the run
/// succeeded, drew, and produced no window.
///
/// It is the ambient document because that is what a caller holding no
/// particular one means. It is also what makes a `document` handle SURVIVE a
/// reset: `dom::reset` clears the map while `next_id` keeps climbing, so a
/// handle that captured id 1 names a document that no longer exists, and every
/// call on it goes quiet. The global `document` is bound once, before the
/// program runs, and must keep meaning the document the program is IN.
fn doc_arg(args: &[Value], idx: usize) -> DocumentId {
    let named = match args.get(idx) {
        Some(Value::Object(o)) => o
            .lock()
            .unwrap()
            .properties
            .get("__document")
            .map(|v| v.as_f64() as DocumentId)
            .unwrap_or(0),
        Some(v) => v.as_f64() as DocumentId,
        None => 0,
    };
    if named == 0 { active_document() } else { named }
}

/// Wrap a document as the object guest code holds.
///
/// The `Document` counterpart to [`element`], and the same three facts:
/// `__node` is the root, `__type` is what it is, `__document` is which tree.
/// A document is its own `ownerDocument`, so the last two are about the same
/// thing — which is what lets one decoder ([`doc_arg`]) read both handles.
pub fn document_handle(document: DocumentId) -> Value {
    let mut o = Object::new_typed(live_type_ids().document);
    o.properties
        .insert("__node".into(), Value::F64(DOCUMENT as f64));
    o.properties
        .insert("__type".into(), Value::String(Arc::from("Document")));
    o.properties
        .insert("__document".into(), Value::F64(document as f64));
    // `document.body` is an IDL ATTRIBUTE (HTML §3.1.1), not a method, and the
    // TypeRegistry vtable holds methods only — so it is a property on the
    // object, which is exactly how `dom_parser` carries `tagName`/`childNodes`
    // and how a plain `Op::STRUCT_GET` reaches it.
    //
    // Without this, `document.body.appendChild(node)` reads `undefined`, calls
    // a method on it, and inserts NOTHING while raising nothing — the page just
    // comes out empty. The `web:html:body` host fn stays: it is the same fact
    // for a caller that imports rather than dispatches.
    o.properties.insert("body".into(), element(document, DOCUMENT));
    Value::Object(vybe_runtime::heap::alloc(o))
}

/// An element reference — the object `createElement` handed back, or a bare
/// handle.
///
/// An `Element` IS an object in the DOM, not an integer: guest code stamps
/// class identity onto it (`__type`, `__types`) and calls methods on it, and
/// a number can carry none of that. The node id travels inside as `__node`.
fn node_arg(args: &[Value], idx: usize) -> NodeId {
    match args.get(idx) {
        Some(Value::Object(o)) => o
            .lock()
            .unwrap()
            .properties
            .get("__node")
            .map(|v| v.as_f64() as NodeId)
            .unwrap_or(DOCUMENT),
        Some(v) => v.as_f64() as NodeId,
        None => DOCUMENT,
    }
}

/// The `TypeDef` ids the LIVE document's handles are stamped with, so a
/// spec-shaped `document.createElement(…)` / `elem.setAttribute(…)` resolves
/// through the TypeRegistry vtable.
///
/// They are set once, from [`crate::builtin_types::register_types`], because
/// the ids only exist after registration. Zero until then, which is `Object` —
/// no methods, so a call fails to resolve rather than reaching the wrong tree.
///
/// **Why these are not the `Element`/`Document` ids.** Those belong to
/// `web:dom-parser`'s trees, whose methods walk detached `Value::Object` nodes;
/// the live document's methods go to `web:dom` and walk `vybe_widgets::dom`.
/// One name cannot carry two implementations, so the live handles are the HTML
/// Standard's own `HTMLDocument`/`HTMLElement`. That the two exist at all is
/// the open item — see the two-DOMs note in the crate docs.
#[derive(Default, Clone, Copy)]
pub struct LiveTypeIds {
    pub document: usize,
    pub element: usize,
}

static LIVE_TYPE_IDS: Mutex<LiveTypeIds> = Mutex::new(LiveTypeIds {
    document: 0,
    element: 0,
});

pub fn set_live_type_ids(ids: LiveTypeIds) {
    *LIVE_TYPE_IDS.lock().unwrap() = ids;
}

fn live_type_ids() -> LiveTypeIds {
    *LIVE_TYPE_IDS.lock().unwrap()
}

/// Wrap a node as the element object guest code holds.
///
/// Three things travel in the handle, and each one answers a question a bare
/// node id cannot:
///
/// - `__node` — which node. The identity.
/// - `__type` — WHAT it is, so `elem.setAttribute(…)` can dispatch through the
///   `Element` vtable. Without it there is no type to resolve a method on, and
///   spec-shaped calls have to be spelled as free functions instead.
/// - `__document` — `Node.ownerDocument` (DOM §4.4). A method called ON an
///   element receives only the element, so the document has to be reachable
///   FROM it or a receiver-shaped call cannot be answered at all.
fn element(document: DocumentId, node: NodeId) -> Value {
    let mut o = Object::new_typed(live_type_ids().element);
    o.properties
        .insert("__node".into(), Value::F64(node as f64));
    o.properties
        .insert("__type".into(), Value::String(Arc::from("Element")));
    o.properties
        .insert("__document".into(), Value::F64(document as f64));
    Value::Object(vybe_runtime::heap::alloc(o))
}

fn str_arg(args: &[Value], idx: usize) -> String {
    args.get(idx)
        .map(|v| format!("{}", v))
        .filter(|s| s != "null" && s != "undefined")
        .unwrap_or_default()
}

fn num_arg(args: &[Value], idx: usize) -> f64 {
    args.get(idx).map(|v| v.as_f64()).unwrap_or(0.0)
}

/// JS truthiness. `Value::as_bool()` is strict about `Bool(true)`, so a guest
/// that computes a flag numerically would otherwise read as false — that bug
/// cost real time on the SDL modifier keys.
fn truthy(v: Option<&Value>) -> bool {
    match v {
        Some(Value::Bool(b)) => *b,
        Some(Value::I32(n)) => *n != 0,
        Some(Value::I64(n)) => *n != 0,
        Some(Value::F64(f)) => *f != 0.0 && !f.is_nan(),
        Some(Value::String(s)) => !s.is_empty() && s.as_ref() != "false",
        Some(Value::Null) | Some(Value::Undefined) | None => false,
        Some(_) => true,
    }
}

// ── result encoding ─────────────────────────────────────────────────────

fn as_node(document: DocumentId, v: DomValue) -> Value {
    match v {
        DomValue::Node(n) => element(document, n),
        _ => Value::Null,
    }
}

fn as_text(v: DomValue) -> Value {
    match v {
        DomValue::Text(s) => Value::String(s.into()),
        _ => Value::String("".into()),
    }
}

fn as_bool(v: DomValue) -> Value {
    Value::Bool(matches!(v, DomValue::Bool(true)))
}

/// A numeric IDL attribute. `selectedIndex` is a `long`, so this answers an
/// integer — `-1` is the IDL's own "nothing selected", and it is also the right
/// answer for an element that has no selection concept at all.
fn as_number(v: DomValue) -> Value {
    match v {
        DomValue::Number(n) => Value::I32(n as i32),
        _ => Value::I32(-1),
    }
}

/// The DOM's own resource types, in Component Model terms.
///
/// A node is a RESOURCE: the host owns the tree, the guest holds a handle. That
/// is what `component_model.rs` describes ("a GUI control is a resource: the
/// host manages the actual widget, the guest holds a handle"), and it is what
/// makes `append-child` a method on `node` instead of a free function that
/// happens to take a node-shaped `Value` first.
///
/// Every DOM operation BORROWS its handles — `appendChild` neither consumes its
/// parent nor its child — which is why `borrows_self` is true and the params
/// are `Borrow`, not `Own`.
const NODE: &str = "node";
const DOCUMENT_RES: &str = "document";

/// `results` is the COMPONENT MODEL result — a void operation declares none.
///
/// That is not the same thing as what the VM does: every host call leaves one
/// value on the stack whatever it returns, which is why the `gui::emit_*`
/// helpers drop it. A setter declaring `vec![]` while its closure answers
/// `Value::Null` is those two conventions meeting, not a mismatch.
fn node_method(name: &str, params: Vec<ValType>, results: Vec<ValType>) -> FuncSig {
    FuncSig {
        name: name.to_string(),
        params,
        results,
    }
}

/// A borrowed node handle — the shape of nearly every DOM parameter.
fn node() -> ValType {
    ValType::Borrow(NODE.to_string())
}

/// A borrowed document handle. Every operation here takes one first: the
/// document owns the tree, so it is the receiver and the node is an argument.
fn doc() -> ValType {
    ValType::Borrow(DOCUMENT_RES.to_string())
}

/// Register a `web:dom` function WITH its signature, in one call.
///
/// Declaration and closure together, deliberately: a signature written beside
/// the registration — or in a table after it — is a second statement of one
/// fact, and the two drift. This is the same reason a control's element and
/// its role live on the type rather than in a lookup elsewhere.
///
/// `kebab` is the Component Model spelling (`append-child`); the registry key
/// stays the camelCase name the emitters already import.
/// What argument 0 is.
enum Receiver {
    /// An element handle — `(node, …)`, missing the document in front.
    Element(DocumentId),
    /// A document handle — already in position, just wrapped.
    Document(DocumentId),
}

/// Which one it is comes from `__type`, NOT from whether `__node` is present.
/// A `Document` handle carries `__node` too — the document's own root — because
/// a document IS a node (DOM §4.5), so presence would classify every document
/// as an element and splice a second document in front of it.
fn receiver(arg: Option<&Value>) -> Option<Receiver> {
    let Some(Value::Object(obj)) = arg else {
        return None;
    };
    let o = obj.lock().unwrap();
    let document = o.properties.get("__document")?.as_f64() as DocumentId;
    Some(match o.properties.get("__type").map(|v| format!("{v}")) {
        Some(t) if t == "Document" => Receiver::Document(document),
        _ => Receiver::Element(document),
    })
}

/// Accept the call in the shape the SPEC writes it, not only the shape the
/// emitters import it.
///
/// WHATWG spells every one of these as a method on a node —
/// `parent.appendChild(child)`, `elem.setAttribute(n, v)`,
/// `document.createElement(tag)` — so dispatching through a type vtable hands
/// the receiver in as argument 0. The closures below all read a DOCUMENT there,
/// because `(doc, node, …)` is the form the emitters have imported from the
/// start. Both forms are the same call; the difference is only where the
/// document comes from.
///
/// So it is expanded once, here, rather than thirty closures each learning to
/// recognise two shapes:
///
/// - an ELEMENT receiver carries `ownerDocument` (DOM §4.4) and is spliced in
///   front, giving `(doc, node, …)`
/// - a DOCUMENT receiver is already in position and only unwraps to its id
/// - anything else is already positional and is passed through untouched
///
/// The alternative was a second, receiver-shaped registration per function —
/// one fact stated twice, which is how the two spellings come to disagree.
fn with_receiver(
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) -> Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync> {
    Box::new(move |ctx: &mut HostContext, args: &[Value]| match receiver(args.first()) {
        Some(Receiver::Element(document)) => {
            let mut expanded = Vec::with_capacity(args.len() + 1);
            expanded.push(Value::F64(document as f64));
            expanded.extend_from_slice(args);
            call(ctx, &expanded)
        }
        Some(Receiver::Document(document)) => {
            let mut positional = args.to_vec();
            positional[0] = Value::F64(document as f64);
            call(ctx, &positional)
        }
        None => call(ctx, args),
    })
}

fn dom_fn(
    vm: &mut VM,
    name: &str,
    kebab: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(
        HostFnDecl::new("web:dom", name, with_receiver(call))
            .with_sig(node_method(kebab, params, results))
            .method_on(DOCUMENT_RES),
    );
}

/// The same, for `web:html` — the HTML element IDL rather than the DOM core.
///
/// The split is the spec's own: `focus()` and `value` are HTMLElement members,
/// `appendChild` is a Node one. Naming the wrong module is not a silent miss
/// but an `Unresolved import` at run time, so the two helpers keep the module
/// out of every call site rather than repeating it 23 times.
fn html_fn(
    vm: &mut VM,
    name: &str,
    kebab: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(
        HostFnDecl::new("web:html", name, with_receiver(call))
            .with_sig(node_method(kebab, params, results))
            .method_on(DOCUMENT_RES),
    );
}

/// And for `web:cssom` — `CSSStyleDeclaration`, whose two operations take a
/// property NAME and a value, both `string`. That is CSSOM's own typing: a
/// declaration's value is text until a property parses it, which is exactly
/// why `vybe_widgets`' `Style` stores declarations verbatim and `CssProperties`
/// is the typed view beside it.
fn css_fn(
    vm: &mut VM,
    name: &str,
    kebab: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(
        HostFnDecl::new("web:cssom", name, with_receiver(call))
            .with_sig(node_method(kebab, params, results))
            .method_on(DOCUMENT_RES),
    );
}

pub fn register(vm: &mut VM) {
    // ── Document ────────────────────────────────────────────────────────
    //
    // Declared through `HostFnDecl`: same closure, same behaviour, plus the
    // signature the registry has never carried. An undeclared registration is
    // untouched — this is a per-function migration, not a flag day.
    vm.register_host(
        HostFnDecl::new(
            "web:dom",
            "createElement",
            with_receiver(Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                as_node(
                    doc_arg(args, 0),
                    apply(
                        doc_arg(args, 0),
                        DomOp::CreateElement {
                            tag: str_arg(args, 1),
                            input_type: str_arg(args, 2),
                        },
                    ),
                )
            })),
        )
        .with_sig(node_method(
            "create-element",
            vec![
                ValType::Borrow(DOCUMENT_RES.to_string()),
                ValType::String,
                ValType::String,
            ],
            vec![ValType::Own(NODE.to_string())],
        ))
        .method_on(DOCUMENT_RES),
    );
    // `nodeType` / `nodeName` / `nodeValue` / `parentNode` / `childNodes` —
    // the read side of a node.
    //
    // **Operations, not properties stamped on a handle.** A handle is minted
    // once and a tree changes: `childNodes` copied onto it is right when taken
    // and wrong immediately after the next `appendChild`. Immutable facts —
    // `ownerDocument`, and the node id itself — may ride on the handle; a live
    // collection may not, and that distinction is the whole reason the
    // property-bag tree could be folded in without a second storage appearing.
    dom_fn(
        vm,
        "nodeType",
        "node-type",
        vec![doc(), node()],
        vec![ValType::I32],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match apply(doc_arg(args, 0), DomOp::NodeType(node_arg(args, 1))) {
                DomValue::Number(n) => Value::I32(n as i32),
                // `0` is not a nodeType, which is what makes it a usable
                // answer for "no such node".
                _ => Value::I32(0),
            }
        }),
    );
    dom_fn(
        vm,
        "nodeName",
        "node-name",
        vec![doc(), node()],
        vec![ValType::String],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_text(apply(doc_arg(args, 0), DomOp::NodeName(node_arg(args, 1))))
        }),
    );
    dom_fn(
        vm,
        "nodeValue",
        "node-value",
        vec![doc(), node()],
        // `null` for an element, per spec — distinct from the `""` an empty
        // comment answers.
        vec![ValType::Option(Box::new(ValType::String))],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match apply(doc_arg(args, 0), DomOp::NodeValue(node_arg(args, 1))) {
                DomValue::Text(v) => Value::String(v.into()),
                _ => Value::Null,
            }
        }),
    );
    dom_fn(
        vm,
        "parentNode",
        "parent-node",
        vec![doc(), node()],
        vec![ValType::Option(Box::new(node()))],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let document = doc_arg(args, 0);
            match apply(document, DomOp::ParentNode(node_arg(args, 1))) {
                DomValue::Node(n) => element(document, n),
                _ => Value::Null,
            }
        }),
    );
    // `Node.firstChild` — DOM §4.4. The first child node, or null.
    //
    // Built on `ChildNodes` rather than a new `DomOp`: the tree already
    // answers "what are this node's children", and "the first one" is a
    // question about that answer, not a second traversal. Adding an op would
    // have put the same walk in two places.
    //
    // `firstChild`, NOT `firstElementChild`: every node type counts, text and
    // comments included. That distinction is the whole reason both members
    // exist in the spec, and picking the wrong one silently skips text.
    //
    // The caller that needed this is `send_to_back` — z-order is document
    // order, so sending a control to the back is `insertBefore` against the
    // parent's current first child. Without it that lowering imported a
    // function nothing registered.
    dom_fn(
        vm,
        "firstChild",
        "first-child",
        vec![doc(), node()],
        vec![ValType::Option(Box::new(node()))],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let document = doc_arg(args, 0);
            match apply(document, DomOp::ChildNodes(node_arg(args, 1))) {
                DomValue::Nodes(ns) => match ns.first() {
                    Some(n) => element(document, *n),
                    None => Value::Null,
                },
                _ => Value::Null,
            }
        }),
    );
    // `Node.lastChild` — DOM §4.4, the mirror of `firstChild`.
    dom_fn(
        vm,
        "lastChild",
        "last-child",
        vec![doc(), node()],
        vec![ValType::Option(Box::new(node()))],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let document = doc_arg(args, 0);
            match apply(document, DomOp::ChildNodes(node_arg(args, 1))) {
                DomValue::Nodes(ns) => match ns.last() {
                    Some(n) => element(document, *n),
                    None => Value::Null,
                },
                _ => Value::Null,
            }
        }),
    );
    // `Node.nextSibling` / `Node.previousSibling` — DOM §4.4.
    //
    // Derived from the PARENT's child list rather than stored links: the tree
    // answers "who are my children" and "who is my parent", and a sibling is
    // the neighbour of this node in that list. A node with no parent has no
    // siblings, which falls out of `ParentNode` answering null.
    for (name, kebab, forward) in [
        ("nextSibling", "next-sibling", true),
        ("previousSibling", "previous-sibling", false),
    ] {
        dom_fn(
            vm,
            name,
            kebab,
            vec![doc(), node()],
            vec![ValType::Option(Box::new(node()))],
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let document = doc_arg(args, 0);
                let me = node_arg(args, 1);
                let DomValue::Node(parent) = apply(document, DomOp::ParentNode(me)) else {
                    return Value::Null;
                };
                let DomValue::Nodes(kids) = apply(document, DomOp::ChildNodes(parent)) else {
                    return Value::Null;
                };
                let Some(i) = kids.iter().position(|n| *n == me) else {
                    return Value::Null;
                };
                let neighbour = if forward {
                    kids.get(i + 1).copied()
                } else {
                    i.checked_sub(1).and_then(|j| kids.get(j).copied())
                };
                match neighbour {
                    Some(n) => element(document, n),
                    None => Value::Null,
                }
            }),
        );
    }
    // `ChildNode.remove()` — DOM §4.2.9. Detach this node from its parent.
    //
    // The spec's own definition is "if I have a parent, remove me from it", so
    // this composes `ParentNode` + `RemoveChild` rather than adding an op. A
    // node with no parent removes nothing and does not raise, per spec.
    dom_fn(
        vm,
        "remove",
        "remove",
        vec![doc(), node()],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let document = doc_arg(args, 0);
            let me = node_arg(args, 1);
            if let DomValue::Node(parent) = apply(document, DomOp::ParentNode(me)) {
                apply(
                    document,
                    DomOp::RemoveChild {
                        parent,
                        child: me,
                    },
                );
            }
            Value::Null
        }),
    );
    // `Element.hasAttribute` — DOM §4.9. Presence, not truthiness: an
    // attribute set to the empty string IS present, which is what makes
    // `<input disabled>` work.
    dom_fn(
        vm,
        "hasAttribute",
        "has-attribute",
        vec![doc(), node(), ValType::String],
        vec![ValType::Bool],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let found = matches!(
                apply(
                    doc_arg(args, 0),
                    DomOp::GetAttribute(node_arg(args, 1), str_arg(args, 2)),
                ),
                DomValue::Text(_)
            );
            Value::Bool(found)
        }),
    );
    // `element.getAttributeNames()` — DOM §4.9. The half of the attribute
    // surface that was missing: everything else here is addressed BY NAME, so
    // nothing could ask an element what it has. A diff needs exactly that —
    // without it a reconciler can compare the attributes it already knows to
    // look for and no others.
    dom_fn(
        vm,
        "getAttributeNames",
        "get-attribute-names",
        vec![doc(), node()],
        vec![ValType::List(Box::new(ValType::String))],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match apply(doc_arg(args, 0), DomOp::AttributeNames(node_arg(args, 1))) {
                DomValue::Texts(names) => {
                    let items: Vec<Value> =
                        names.into_iter().map(|n| Value::String(n.into())).collect();
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(items)))
                }
                _ => Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new()))),
            }
        }),
    );
    // ── `Element.classList` — DOM §4.9 / DOMTokenList §7.1 ────────────────
    //
    // FLAT functions, not an object: the same shape `style.setProperty` already
    // takes as `setStyleProperty`. A host function surface has no place to hang
    // a live `DOMTokenList`, and inventing one would be a second way to say
    // what `class` already says.
    //
    // The token list IS the `class` attribute, parsed on demand — so these read
    // and write through `GetAttribute`/`SetAttribute` and stay consistent with
    // anything that touched `class` directly. Serialised back space-separated,
    // in order, which is what the spec's "ordered set serializer" produces.
    //
    // Without these the only way to toggle a class was read the attribute,
    // splice a string, and write it back — in the guest, differently each time.
    fn tokens(document: crate::engine::DocumentId, n: crate::engine::NodeId) -> Vec<String> {
        match apply(document, DomOp::GetAttribute(n, "class".to_string())) {
            DomValue::Text(s) => s.split_whitespace().map(str::to_string).collect(),
            _ => Vec::new(),
        }
    }
    fn write_tokens(
        document: crate::engine::DocumentId,
        n: crate::engine::NodeId,
        list: &[String],
    ) {
        apply(
            document,
            DomOp::SetAttribute(n, "class".to_string(), list.join(" ")),
        );
    }
    dom_fn(
        vm,
        "classListAdd",
        "class-list-add",
        vec![doc(), node(), ValType::String],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let (document, n, token) = (doc_arg(args, 0), node_arg(args, 1), str_arg(args, 2));
            if token.is_empty() {
                return Value::Null;
            }
            let mut list = tokens(document, n);
            // A set: adding a token already present is a no-op, not a
            // duplicate. That is what makes `add` idempotent per spec.
            if !list.iter().any(|t| *t == token) {
                list.push(token);
                write_tokens(document, n, &list);
            }
            Value::Null
        }),
    );
    dom_fn(
        vm,
        "classListRemove",
        "class-list-remove",
        vec![doc(), node(), ValType::String],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let (document, n, token) = (doc_arg(args, 0), node_arg(args, 1), str_arg(args, 2));
            let mut list = tokens(document, n);
            let before = list.len();
            list.retain(|t| *t != token);
            if list.len() != before {
                write_tokens(document, n, &list);
            }
            Value::Null
        }),
    );
    dom_fn(
        vm,
        "classListContains",
        "class-list-contains",
        vec![doc(), node(), ValType::String],
        vec![ValType::Bool],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let (document, n, token) = (doc_arg(args, 0), node_arg(args, 1), str_arg(args, 2));
            Value::Bool(tokens(document, n).iter().any(|t| *t == token))
        }),
    );
    dom_fn(
        vm,
        "classListToggle",
        "class-list-toggle",
        vec![doc(), node(), ValType::String],
        vec![ValType::Bool],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let (document, n, token) = (doc_arg(args, 0), node_arg(args, 1), str_arg(args, 2));
            if token.is_empty() {
                return Value::Bool(false);
            }
            let mut list = tokens(document, n);
            // Returns whether the token is present AFTER the call — the spec's
            // return value, and the reason this is not just add-or-remove.
            let present = list.iter().any(|t| *t == token);
            if present {
                list.retain(|t| *t != token);
            } else {
                list.push(token);
            }
            write_tokens(document, n, &list);
            Value::Bool(!present)
        }),
    );
    // ── Element-only traversal — DOM §4.2.6 `ParentNode` / `NonDocumentTypeChildNode`
    //
    // `children` is NOT `childNodes`: it skips text and comments. That is the
    // whole reason the spec has both, and the difference bites the moment a
    // document has whitespace between tags — which parsed markup always does.
    //
    // ELEMENT_NODE is 1 (DOM §4.4). Filtering on `NodeType` rather than a new
    // op keeps this derived from what the tree already answers.
    fn element_children(
        document: crate::engine::DocumentId,
        n: crate::engine::NodeId,
    ) -> Vec<crate::engine::NodeId> {
        let DomValue::Nodes(kids) = apply(document, DomOp::ChildNodes(n)) else {
            return Vec::new();
        };
        kids.into_iter()
            .filter(|k| matches!(apply(document, DomOp::NodeType(*k)), DomValue::Number(t) if t as i32 == 1))
            .collect()
    }
    dom_fn(
        vm,
        "children",
        "children",
        vec![doc(), node()],
        vec![ValType::List(Box::new(node()))],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let document = doc_arg(args, 0);
            let items: Vec<Value> = element_children(document, node_arg(args, 1))
                .into_iter()
                .map(|n| element(document, n))
                .collect();
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(items)))
        }),
    );
    for (name, kebab, last) in [
        ("firstElementChild", "first-element-child", false),
        ("lastElementChild", "last-element-child", true),
    ] {
        dom_fn(
            vm,
            name,
            kebab,
            vec![doc(), node()],
            vec![ValType::Option(Box::new(node()))],
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let document = doc_arg(args, 0);
                let kids = element_children(document, node_arg(args, 1));
                let pick = if last { kids.last() } else { kids.first() };
                match pick {
                    Some(n) => element(document, *n),
                    None => Value::Null,
                }
            }),
        );
    }
    // `Node.contains(other)` — DOM §4.4. True if `other` is this node or a
    // descendant of it. Inclusive, per spec: a node contains itself.
    dom_fn(
        vm,
        "contains",
        "contains",
        vec![doc(), node(), node()],
        vec![ValType::Bool],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let document = doc_arg(args, 0);
            let (me, other) = (node_arg(args, 1), node_arg(args, 2));
            let mut cursor = Some(other);
            while let Some(n) = cursor {
                if n == me {
                    return Value::Bool(true);
                }
                cursor = match apply(document, DomOp::ParentNode(n)) {
                    DomValue::Node(p) => Some(p),
                    _ => None,
                };
            }
            Value::Bool(false)
        }),
    );
    // `Element.matches(selector)` — DOM §4.9.
    //
    // Derived from `querySelectorAll` + membership rather than a new matcher:
    // the engine already owns selector semantics, and a SECOND implementation
    // is how the two drift. Costs a document-wide query per call, which is the
    // honest trade for having exactly one matcher.
    fn matches_selector(
        document: crate::engine::DocumentId,
        n: crate::engine::NodeId,
        selector: &str,
    ) -> bool {
        match apply(document, DomOp::QuerySelectorAll(selector.to_string())) {
            DomValue::Nodes(ns) => ns.contains(&n),
            _ => false,
        }
    }
    dom_fn(
        vm,
        "matches",
        "matches",
        vec![doc(), node(), ValType::String],
        vec![ValType::Bool],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(matches_selector(
                doc_arg(args, 0),
                node_arg(args, 1),
                &str_arg(args, 2),
            ))
        }),
    );
    // `Element.closest(selector)` — DOM §4.9. This node, then each ancestor,
    // first match wins. Inclusive of self, per spec.
    dom_fn(
        vm,
        "closest",
        "closest",
        vec![doc(), node(), ValType::String],
        vec![ValType::Option(Box::new(node()))],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let document = doc_arg(args, 0);
            let selector = str_arg(args, 2);
            let mut cursor = Some(node_arg(args, 1));
            while let Some(n) = cursor {
                if matches_selector(document, n, &selector) {
                    return element(document, n);
                }
                cursor = match apply(document, DomOp::ParentNode(n)) {
                    DomValue::Node(p) => Some(p),
                    _ => None,
                };
            }
            Value::Null
        }),
    );
    dom_fn(
        vm,
        "childNodes",
        "child-nodes",
        vec![doc(), node()],
        vec![ValType::List(Box::new(node()))],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let document = doc_arg(args, 0);
            let items: Vec<Value> = match apply(document, DomOp::ChildNodes(node_arg(args, 1))) {
                DomValue::Nodes(ns) => ns.into_iter().map(|n| element(document, n)).collect(),
                _ => Vec::new(),
            };
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(items)))
        }),
    );

    // `insertBefore` / `replaceChild` / `cloneNode` — DOM §4.2.3 and §4.4.
    //
    // The three mutations the seam has never had. `appendChild` and
    // `removeChild` can express "in this parent" and not "in this ORDER", so
    // anything building a tree at a position — a diffing renderer, an XML
    // document, `innerHTML` — had no operation to call.
    //
    // Each answers `false`/`null` rather than trapping. `insertBefore` with a
    // reference that is not a child is the spec's `NotFoundError`, and there is
    // no exception channel here: refusing is the safe direction, appending
    // would be a wrong ORDER, which is invisible until someone reads the tree
    // back.
    dom_fn(
        vm,
        "insertBefore",
        "insert-before",
        vec![doc(), node(), node(), node()],
        vec![ValType::Bool],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_bool(apply(
                doc_arg(args, 0),
                DomOp::InsertBefore {
                    parent: node_arg(args, 1),
                    child: node_arg(args, 2),
                    reference: node_arg(args, 3),
                },
            ))
        }),
    );
    dom_fn(
        vm,
        "replaceChild",
        "replace-child",
        vec![doc(), node(), node(), node()],
        vec![ValType::Bool],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_bool(apply(
                doc_arg(args, 0),
                DomOp::ReplaceChild {
                    parent: node_arg(args, 1),
                    new_child: node_arg(args, 2),
                    old_child: node_arg(args, 3),
                },
            ))
        }),
    );
    dom_fn(
        vm,
        "cloneNode",
        "clone-node",
        vec![doc(), node(), ValType::Bool],
        vec![ValType::Own(NODE.to_string())],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let document = doc_arg(args, 0);
            as_node(
                document,
                apply(
                    document,
                    DomOp::CloneNode {
                        node: node_arg(args, 1),
                        // `cloneNode()` with no argument is a SHALLOW clone —
                        // the spec's default is `false`, and `truthy` on an
                        // absent argument is what says so.
                        deep: truthy(args.get(2)),
                    },
                ),
            )
        }),
    );

    // `document.createTextNode(data)` / `document.createComment(data)` — the
    // DOM's other two node factories.
    //
    // The surface had `createElement` and nothing else, so the CONTENT between
    // two elements could not be created at all: a guest could build
    // `<p></p><p></p>` and never `a <b>B</b> c`. Both answer an uninserted
    // node, exactly as `createElement` does — it renders nothing until
    // `appendChild` puts it somewhere.
    dom_fn(
        vm,
        "createTextNode",
        "create-text-node",
        vec![doc(), ValType::String],
        vec![ValType::Own(NODE.to_string())],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let document = doc_arg(args, 0);
            as_node(
                document,
                apply(document, DomOp::CreateTextNode(str_arg(args, 1))),
            )
        }),
    );
    dom_fn(
        vm,
        "createComment",
        "create-comment",
        vec![doc(), ValType::String],
        vec![ValType::Own(NODE.to_string())],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let document = doc_arg(args, 0);
            as_node(
                document,
                apply(document, DomOp::CreateComment(str_arg(args, 1))),
            )
        }),
    );
    // `setAttributeNS` / `getAttributeNS`. The asymmetry is the spec's: the
    // write names the attribute the way it serialises (QUALIFIED) and the read
    // matches namespace + LOCAL name, because that pair is what identifies an
    // attribute and a qualified name is only how it is written down.
    dom_fn(
        vm,
        "setAttributeNS",
        "set-attribute-ns",
        vec![doc(), node(), ValType::String, ValType::String, ValType::String],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::SetAttributeNS {
                    node: node_arg(args, 1),
                    namespace: str_arg(args, 2),
                    qualified_name: str_arg(args, 3),
                    value: str_arg(args, 4),
                },
            );
            Value::Null
        }),
    );
    dom_fn(
        vm,
        "getAttributeNS",
        "get-attribute-ns",
        vec![doc(), node(), ValType::String, ValType::String],
        vec![ValType::Option(Box::new(ValType::String))],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match apply(
                doc_arg(args, 0),
                DomOp::GetAttributeNS {
                    node: node_arg(args, 1),
                    namespace: str_arg(args, 2),
                    local_name: str_arg(args, 3),
                },
            ) {
                DomValue::Text(v) => Value::String(v.into()),
                _ => Value::Null,
            }
        }),
    );

    // `createElementNS` and the three reads that make a namespace visible.
    // `prefix` and `localName` are views of the qualified name rather than
    // stored fields, so the engine derives them and nothing here can disagree
    // with `nodeName`.
    dom_fn(
        vm,
        "createElementNS",
        "create-element-ns",
        vec![doc(), ValType::String, ValType::String, ValType::String],
        vec![ValType::Own(NODE.to_string())],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let document = doc_arg(args, 0);
            as_node(
                document,
                apply(
                    document,
                    DomOp::CreateElementNS {
                        namespace: str_arg(args, 1),
                        qualified_name: str_arg(args, 2),
                        input_type: str_arg(args, 3),
                    },
                ),
            )
        }),
    );
    for (name, kebab) in [
        ("namespaceURI", "namespace-uri"),
        ("prefix", "prefix"),
        ("localName", "local-name"),
    ] {
        dom_fn(
            vm,
            name,
            kebab,
            vec![doc(), node()],
            // `namespaceURI` and `prefix` are nullable; `localName` never is.
            vec![ValType::Option(Box::new(ValType::String))],
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let node = node_arg(args, 1);
                let op = match name {
                    "namespaceURI" => DomOp::NamespaceUri(node),
                    "prefix" => DomOp::Prefix(node),
                    _ => DomOp::LocalName(node),
                };
                match apply(doc_arg(args, 0), op) {
                    DomValue::Text(v) => Value::String(v.into()),
                    _ => Value::Null,
                }
            }),
        );
    }

    // The XML half of the factory set. Both exist here rather than only in
    // `web:dom-parser` because the node they make is a node in the ONE
    // document — the parser's separate tree is what folds into this.
    dom_fn(
        vm,
        "createCDATASection",
        "create-cdata-section",
        vec![doc(), ValType::String],
        vec![ValType::Own(NODE.to_string())],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let document = doc_arg(args, 0);
            as_node(
                document,
                apply(document, DomOp::CreateCDataSection(str_arg(args, 1))),
            )
        }),
    );
    dom_fn(
        vm,
        "createProcessingInstruction",
        "create-processing-instruction",
        vec![doc(), ValType::String, ValType::String],
        vec![ValType::Own(NODE.to_string())],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let document = doc_arg(args, 0);
            as_node(
                document,
                apply(
                    document,
                    DomOp::CreateProcessingInstruction {
                        target: str_arg(args, 1),
                        data: str_arg(args, 2),
                    },
                ),
            )
        }),
    );

    vm.register_host(
        HostFnDecl::new(
            "web:dom",
            "getElementById",
            with_receiver(Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                as_node(
                    doc_arg(args, 0),
                    apply(doc_arg(args, 0), DomOp::GetElementById(str_arg(args, 1))),
                )
            })),
        )
        .with_sig(node_method(
            "get-element-by-id",
            vec![ValType::Borrow(DOCUMENT_RES.to_string()), ValType::String],
            vec![ValType::Option(Box::new(ValType::Borrow(
                NODE.to_string(),
            )))],
        ))
        .method_on(DOCUMENT_RES),
    );
    // `querySelector` / `querySelectorAll` — Selectors API Level 1 over the
    // LIVE document.
    //
    // These did not exist here at all: the only selector engine was
    // `web:dom-parser`'s, wired to a parsed tree that renders nothing, so a
    // page could not ask its own document a question richer than a tag name.
    // The engine below the seam owns the matching now, which is what makes
    // these two host functions pure forwarding like everything else in this
    // file.
    dom_fn(
        vm,
        "querySelector",
        "query-selector",
        vec![doc(), ValType::String],
        vec![node()],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match apply(doc_arg(args, 0), DomOp::QuerySelector(str_arg(args, 1))) {
                DomValue::Node(n) => element(doc_arg(args, 0), n),
                // No match is `null`, per spec — not an empty element.
                _ => Value::Null,
            }
        }),
    );
    dom_fn(
        vm,
        "querySelectorAll",
        "query-selector-all",
        vec![doc(), ValType::String],
        vec![ValType::List(Box::new(node()))],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match apply(doc_arg(args, 0), DomOp::QuerySelectorAll(str_arg(args, 1))) {
                DomValue::Nodes(ns) => {
                    let d = doc_arg(args, 0);
                    let items: Vec<Value> = ns.into_iter().map(|n| element(d, n)).collect();
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(items)))
                }
                _ => Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new()))),
            }
        }),
    );
    // An `HTMLCollection` is a LIST of borrowed nodes: the collection does not
    // own what it names, and neither does the guest that reads it. Nothing in
    // `web:dom` calls this today — the XML path goes to `web:dom-parser`'s own
    // `getElementsByTagName` — so the declaration is the only statement of its
    // shape there has ever been.
    dom_fn(
        vm,
        "getElementsByTagName",
        "get-elements-by-tag-name",
        vec![doc(), ValType::String],
        vec![ValType::List(Box::new(node()))],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match apply(doc_arg(args, 0), DomOp::ElementsByTag(str_arg(args, 1))) {
                DomValue::Nodes(ns) => {
                    let d = doc_arg(args, 0);
                    let items: Vec<Value> = ns.into_iter().map(|n| element(d, n)).collect();
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(items)))
                }
                _ => Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new()))),
            }
        }),
    );
    // `document` — the active document of the current browsing context.
    //
    // In a page this is simply ambient: script always has `document`. A guest
    // that has not opened a window yet still gets one, materialised on first
    // use exactly as a browsing context's initial `about:blank` is. Without
    // it a control could not be created before a window exists, which every
    // form-designer-shaped program does.
    //
    // The one STATIC here: it takes no document because it is how you get one.
    // Every other function in this file is a method whose receiver is the
    // handle this returns, which is why it declares no parameters and they all
    // declare `borrow<document>` first.
    vm.register_host(
        HostFnDecl::new(
            "web:html",
            "activeDocument",
            Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
                document_handle(active_document())
            }),
        )
        .with_sig(node_method(
            "active-document",
            vec![],
            vec![ValType::Borrow(DOCUMENT_RES.to_string())],
        ))
        .resource_member(ResourceBinding {
            resource: DOCUMENT_RES.to_string(),
            kind: ResourceMemberKind::Static,
            // A static has no self to borrow. Every method below says `true`
            // because it holds a `borrow<document>`; this one is how you GET
            // that handle, so there is nothing yet to hold.
            borrows_self: false,
        }),
    );

    // `document.body` — the document element every control hangs off.
    //
    // The parameter is declared because `body` IS a member of a document and
    // the caller passes one. The closure ignores it and answers the one
    // `DOCUMENT` root, so this is single-document today; the declaration is
    // what will make that visible the day a second document exists, rather
    // than the argument silently going nowhere as it does now.
    html_fn(
        vm,
        "body",
        "body",
        vec![doc()],
        vec![node()],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            element(doc_arg(args, 0), DOCUMENT)
        }),
    );
    // `document.defaultView` (HTML §3.1.1) — the WindowProxy of this document's
    // browsing context, or null if it has none.
    //
    // **This is how a guest names its own window.** In a browser you never ask
    // for the current window, you ARE in it: `window`/`self`/`globalThis` are
    // the global object. A guest here holds no such global, so the standard
    // document→window accessor is the door, and it is the exact inverse of the
    // `window.document` already registered next door.
    //
    // A document with no context answers NULL rather than inventing one — a
    // `DOMParser` document genuinely has no `defaultView`, and creating a
    // browsing context for it would put a tab on screen for a parsed string.
    html_fn(
        vm,
        "defaultView",
        "default-view",
        vec![doc()],
        vec![ValType::Borrow("window".to_string())],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match crate::engine::window(crate::engine::WindowOp::DefaultView(doc_arg(args, 0))) {
                crate::engine::WindowValue::Window(w) => Value::F64(w as f64),
                _ => Value::Null,
            }
        }),
    );
    html_fn(
        vm,
        "title",
        "title",
        vec![doc()],
        vec![ValType::String],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_text(apply(doc_arg(args, 0), DomOp::Title))
        }),
    );
    html_fn(
        vm,
        "setTitle",
        "set-title",
        vec![doc(), ValType::String],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(doc_arg(args, 0), DomOp::SetTitle(str_arg(args, 1)));
            Value::Null
        }),
    );

    // ── Node ────────────────────────────────────────────────────────────
    dom_fn(
        vm,
        "appendChild",
        "append-child",
        vec![doc(), node(), node()],
        vec![node()],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::AppendChild {
                    parent: node_arg(args, 1),
                    child: node_arg(args, 2),
                },
            );
            // `appendChild` returns the appended child.
            args.get(2).cloned().unwrap_or(Value::Null)
        }),
    );
    dom_fn(
        vm,
        "removeChild",
        "remove-child",
        vec![doc(), node(), node()],
        vec![node()],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::RemoveChild {
                    parent: node_arg(args, 1),
                    child: node_arg(args, 2),
                },
            );
            args.get(2).cloned().unwrap_or(Value::Null)
        }),
    );
    dom_fn(
        vm,
        "isConnected",
        "is-connected",
        vec![doc(), node()],
        vec![ValType::Bool],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_bool(apply(
                doc_arg(args, 0),
                DomOp::IsConnected(node_arg(args, 1)),
            ))
        }),
    );
    dom_fn(
        vm,
        "setTextContent",
        "set-text-content",
        vec![doc(), node(), ValType::String],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::SetTextContent(node_arg(args, 1), str_arg(args, 2)),
            );
            Value::Null
        }),
    );
    dom_fn(
        vm,
        "textContent",
        "text-content",
        vec![doc(), node()],
        vec![ValType::String],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_text(apply(
                doc_arg(args, 0),
                DomOp::TextContent(node_arg(args, 1)),
            ))
        }),
    );

    // ── Element: attributes ─────────────────────────────────────────────
    dom_fn(
        vm,
        "setAttribute",
        "set-attribute",
        vec![doc(), node(), ValType::String, ValType::String],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::SetAttribute(node_arg(args, 1), str_arg(args, 2), str_arg(args, 3)),
            );
            Value::Null
        }),
    );
    dom_fn(
        vm,
        "getAttribute",
        "get-attribute",
        vec![doc(), node(), ValType::String],
        // An absent attribute is `null`, per spec — `option<string>`, not a
        // string that happens to be empty. The declaration says so.
        vec![ValType::Option(Box::new(ValType::String))],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match apply(
                doc_arg(args, 0),
                DomOp::GetAttribute(node_arg(args, 1), str_arg(args, 2)),
            ) {
                DomValue::Text(s) => Value::String(s.into()),
                _ => Value::Null,
            }
        }),
    );
    // `element.toggleAttribute(qualifiedName, force)` — DOM Standard.
    //
    // Boolean content attributes are true by PRESENCE: `disabled=""` disables
    // and the attribute must be REMOVED to enable. A plain `setAttribute`
    // would disable a control when you enabled it, so the spec's own
    // add-or-remove primitive is the correct one rather than two calls and a
    // branch in the emitter.
    dom_fn(
        vm,
        "toggleAttribute",
        "toggle-attribute",
        vec![doc(), node(), ValType::String, ValType::Bool],
        vec![ValType::Bool],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(doc) = doc_arg_id(args, 0) else {
                return Value::Bool(false);
            };
            let (node, name) = (node_arg(args, 1), str_arg(args, 2));
            let force = truthy(args.get(3));
            if force {
                apply(doc, DomOp::SetAttribute(node, name, String::new()));
            } else {
                apply(doc, DomOp::RemoveAttribute(node, name));
            }
            Value::Bool(force)
        }),
    );
    dom_fn(
        vm,
        "removeAttribute",
        "remove-attribute",
        vec![doc(), node(), ValType::String],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::RemoveAttribute(node_arg(args, 1), str_arg(args, 2)),
            );
            Value::Null
        }),
    );

    // ── CSSStyleDeclaration ─────────────────────────────────────────────
    css_fn(
        vm,
        "setStyleProperty",
        "set-property",
        vec![doc(), node(), ValType::String, ValType::String],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::SetStyleProperty(node_arg(args, 1), str_arg(args, 2), str_arg(args, 3)),
            );
            Value::Null
        }),
    );
    // `CSSStyleDeclaration.removeProperty(name)` — CSSOM §6.7.1.
    //
    // Returns the OLD value, which is what makes it more than a setter: the
    // spec has it answer the removed declaration's value (or `""` when there
    // was none), so a caller can restore it.
    //
    // Removal is spelled as setting the empty string, because the declaration
    // store treats "" as absent — the same convention `setProperty(name, "")`
    // has in a browser. That keeps one storage rule instead of adding a
    // `RemoveStyleProperty` op that would have to agree with it.
    css_fn(
        vm,
        "removeStyleProperty",
        "remove-property",
        vec![doc(), node(), ValType::String],
        vec![ValType::String],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let (document, n, prop) = (doc_arg(args, 0), node_arg(args, 1), str_arg(args, 2));
            let old = match apply(document, DomOp::GetStyleProperty(n, prop.clone())) {
                DomValue::Text(s) => s,
                _ => String::new(),
            };
            apply(
                document,
                DomOp::SetStyleProperty(n, prop, String::new()),
            );
            Value::String(std::sync::Arc::from(old.as_str()))
        }),
    );
    // `getPropertyValue` answers `""` for a property that is not set — CSSOM
    // says so outright, and it is the one place a DOM read is NOT nullable.
    // `getAttribute` next door returns `option<string>` for exactly the same
    // question about an attribute; the two differ and the declarations say so.
    css_fn(
        vm,
        "getStyleProperty",
        "get-property-value",
        vec![doc(), node(), ValType::String],
        vec![ValType::String],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_text(apply(
                doc_arg(args, 0),
                DomOp::GetStyleProperty(node_arg(args, 1), str_arg(args, 2)),
            ))
        }),
    );

    // `getComputedStyle(el).getPropertyValue(p)` — the RESOLVED value.
    //
    // The sibling of the declaration read above, and a different question:
    // `getPropertyValue` on `element.style` serializes what was DECLARED, while
    // this one answers in used units after cascade and layout. `setProperty
    // ("top", "1em")` reads back `"1em"` there and `"16px"` here.
    //
    // Registered as its own name because a toolkit geometry read
    // (`Control.Left`) means THIS one — it wants the laid-out pixel — while a
    // frontend round-tripping a stylesheet means the other. One function
    // serving both is how the two engines came to disagree.
    css_fn(
        vm,
        "getComputedStyleProperty",
        "get-computed-property-value",
        vec![doc(), node(), ValType::String],
        vec![ValType::String],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_text(apply(
                doc_arg(args, 0),
                DomOp::ComputedStyleProperty(node_arg(args, 1), str_arg(args, 2)),
            ))
        }),
    );

    // ── HTMLInputElement / HTMLSelectElement IDL ────────────────────────
    // `value` is a DOMString in the IDL whatever the input's type is — a
    // number field's value is the TEXT the user typed, which is why an empty
    // one answers `""` and not zero. Declaring it `string` is what stops a
    // frontend assuming the host coerced.
    html_fn(
        vm,
        "setValue",
        "set-value",
        vec![doc(), node(), ValType::String],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::SetValue(node_arg(args, 1), str_arg(args, 2)),
            );
            Value::Null
        }),
    );
    html_fn(
        vm,
        "value",
        "value",
        vec![doc(), node()],
        vec![ValType::String],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_text(apply(doc_arg(args, 0), DomOp::Value(node_arg(args, 1))))
        }),
    );
    html_fn(
        vm,
        "setChecked",
        "set-checked",
        vec![doc(), node(), ValType::Bool],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::SetChecked(node_arg(args, 1), truthy(args.get(2))),
            );
            Value::Null
        }),
    );
    html_fn(
        vm,
        "checked",
        "checked",
        vec![doc(), node()],
        vec![ValType::Bool],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_bool(apply(doc_arg(args, 0), DomOp::Checked(node_arg(args, 1))))
        }),
    );
    html_fn(
        vm,
        "focus",
        "focus",
        vec![doc(), node()],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(doc_arg(args, 0), DomOp::Focus(node_arg(args, 1)));
            Value::Null
        }),
    );

    // `canvas.width` / `canvas.height` — HTMLCanvasElement's own IDL
    // attributes, the READ side of the content attributes `setAttribute`
    // already writes (HTML §4.12.5).
    //
    // Without them the one operation every painter needs cannot be spelled.
    // "Fill the surface with a colour" is `fillRect(0, 0, canvas.width,
    // canvas.height)` — the canvas API has no clear-to-colour and does not
    // need one — so a guest that could size a buffer and never ask its size
    // had to give up the colour. `.NET`'s `Graphics.Clear(color)` was
    // dropping its argument for exactly this reason.
    //
    // Spelled `canvasWidth`/`canvasHeight` rather than `width`/`height`
    // because a bare `width` in this module would read as the BOX's, and the
    // whole point of these two is that they are not it: a 640x480 buffer
    // displayed in a 320x240 box is the ordinary way to draw at double
    // density.
    for (name, kebab, horizontal) in [
        ("canvasWidth", "canvas-width", true),
        ("canvasHeight", "canvas-height", false),
    ] {
        html_fn(
            vm,
            name,
            kebab,
            vec![doc(), node()],
            vec![ValType::F64],
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                match apply(doc_arg(args, 0), DomOp::CanvasSize(node_arg(args, 1))) {
                    DomValue::Pair(w, h) => Value::F64(if horizontal { w } else { h }),
                    // The spec's missing-value defaults, which is what an
                    // element with no bitmap answers rather than zero — a
                    // zero would make `fillRect` over it silently draw
                    // nothing.
                    _ => Value::F64(if horizontal { 300.0 } else { 150.0 }),
                }
            }),
        );
    }

    // `select.selectedIndex` — `-1` when nothing is selected, which is the
    // IDL's own answer and what every caller tests against.
    //
    // `i32`, SIGNED, and that is the whole reason it is not a `u32` index:
    // `-1` is a legal answer and every caller tests for it. A declaration that
    // said `u32` would have made "nothing selected" unrepresentable.
    html_fn(
        vm,
        "selectedIndex",
        "selected-index",
        vec![doc(), node()],
        vec![ValType::I32],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_number(apply(
                doc_arg(args, 0),
                DomOp::SelectedIndex(node_arg(args, 1)),
            ))
        }),
    );
    html_fn(
        vm,
        "setSelectedIndex",
        "set-selected-index",
        vec![doc(), node(), ValType::I32],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::SetSelectedIndex(node_arg(args, 1), num_arg(args, 2) as i32),
            );
            Value::Null
        }),
    );

    // `select.options[i].text` — the option list read and written BY INDEX,
    // which is what a toolkit's `Items[i]` is. `value` cannot stand in: it
    // answers only for the SELECTED option, so every other row is unreachable
    // through it.
    //
    // The index reads `i32` like `selectedIndex`, but it means something
    // narrower: a POSITION, with no negative answer. `ValType` has no unsigned
    // integer to say that in, so the closures clamp (`.max(0.0)`) and this
    // comment is the only place the difference is stated — worth knowing when
    // the type set grows.
    html_fn(
        vm,
        "itemText",
        "item-text",
        vec![doc(), node(), ValType::I32],
        vec![ValType::String],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_text(apply(
                doc_arg(args, 0),
                DomOp::ItemText(node_arg(args, 1), num_arg(args, 2).max(0.0) as usize),
            ))
        }),
    );
    html_fn(
        vm,
        "setItemText",
        "set-item-text",
        vec![doc(), node(), ValType::I32, ValType::String],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::SetItemText(
                    node_arg(args, 1),
                    num_arg(args, 2).max(0.0) as usize,
                    str_arg(args, 3),
                ),
            );
            Value::Null
        }),
    );

    // `select.add(option)` / `select.remove(index)` / `select.length = 0`.
    html_fn(
        vm,
        "addItem",
        "add",
        vec![doc(), node(), ValType::String],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::AddItem(node_arg(args, 1), str_arg(args, 2)),
            );
            Value::Null
        }),
    );
    html_fn(
        vm,
        "removeItem",
        "remove",
        vec![doc(), node(), ValType::I32],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::RemoveItem(node_arg(args, 1), num_arg(args, 2) as usize),
            );
            Value::Null
        }),
    );
    html_fn(
        vm,
        "clearItems",
        "clear-items",
        vec![doc(), node()],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(doc_arg(args, 0), DomOp::ClearItems(node_arg(args, 1)));
            Value::Null
        }),
    );

    // ── HTMLDialogElement ───────────────────────────────────────────────
    //
    // `show()` and `showModal()` RETURN IMMEDIATELY — that is the spec, and
    // it is the one place a toolkit's habits and HTML's genuinely differ. A
    // VCL/WinForms `ShowModal` blocks until the dialog closes; HTML's does
    // not, and you learn the outcome from the `close` event or by polling
    // `open`. Nothing here blocks to make a frontend's life easier: a
    // language that needs blocking spells the wait in its own adapter, on top
    // of `open`, exactly as `FormatDateTime` spells format letters on top of
    // a shared date.
    html_fn(
        vm,
        "show",
        "show",
        vec![doc(), node()],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::ShowDialog {
                    node: node_arg(args, 1),
                    modal: false,
                },
            );
            Value::Null
        }),
    );
    html_fn(
        vm,
        "showModal",
        "show-modal",
        vec![doc(), node()],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::ShowDialog {
                    node: node_arg(args, 1),
                    modal: true,
                },
            );
            Value::Null
        }),
    );

    // `close(returnValue?)` — the optional argument SETS `returnValue`, which
    // is why it is one function and not two.
    //
    // DECLARED WITH TWO PARAMETERS, not three, because two is what the call
    // sites are: every route to a control verb is `emit_gui_control_method`,
    // which emits `(document, control)` and has nowhere to put a third
    // operand. The `args.len() > 2` branch below is therefore unreachable —
    // dead until an emitter grows a way to pass it.
    //
    // Declaring three to match the IDL would make the check fire on every
    // dialog close, which is the failure mode a declaration exists to prevent,
    // not to cause. The Component Model has no optional parameter — an IDL
    // `optional` is `option<T>`, still positional — so honouring the IDL means
    // the CALL SITE passing `none`, which is emitter work and not this task's.
    html_fn(
        vm,
        "close",
        "close",
        vec![doc(), node()],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let (document, node) = (doc_arg(args, 0), node_arg(args, 1));
            if args.len() > 2 {
                return_values()
                    .lock()
                    .unwrap()
                    .0
                    .insert((document, node), str_arg(args, 2));
            }
            apply(document, DomOp::CloseDialog(node));
            Value::Null
        }),
    );

    // `dialog.open` — the reflected attribute, so it answers off the tree.
    html_fn(
        vm,
        "open",
        "open",
        vec![doc(), node()],
        vec![ValType::Bool],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_bool(apply(
                doc_arg(args, 0),
                DomOp::DialogOpen(node_arg(args, 1)),
            ))
        }),
    );

    // `dialog.returnValue` — a DOMString, `""` before anything sets it.
    html_fn(
        vm,
        "returnValue",
        "return-value",
        vec![doc(), node()],
        vec![ValType::String],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let key = (doc_arg(args, 0), node_arg(args, 1));
            let value = return_values()
                .lock()
                .unwrap()
                .0
                .get(&key)
                .cloned()
                .unwrap_or_default();
            Value::String(value.into())
        }),
    );
    // Three parameters, and NOTHING reaches it: a control verb emits two, and
    // `returnvalue` is not a `property_op` role, so a `ReturnValue := 'ok'`
    // write falls to the attribute catch-all and never arrives here. Declared
    // truthfully rather than trimmed to two to look reachable — the mismatch
    // is the finding.
    html_fn(
        vm,
        "setReturnValue",
        "set-return-value",
        vec![doc(), node(), ValType::String],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let key = (doc_arg(args, 0), node_arg(args, 1));
            return_values()
                .lock()
                .unwrap()
                .0
                .insert(key, str_arg(args, 2));
            Value::Null
        }),
    );

    // ── HTMLInputElement.showPicker() ───────────────────────────────────
    //
    // Void, per the IDL. What the user chose lands in `value`, which is where
    // a browser puts it too — so a caller reads the result the same way it
    // reads anything the user typed.
    html_fn(
        vm,
        "showPicker",
        "show-picker",
        vec![doc(), node()],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(doc_arg(args, 0), DomOp::ShowPicker(node_arg(args, 1)));
            Value::Null
        }),
    );

    // ── EventTarget ─────────────────────────────────────────────────────
    //
    // The listener is `Any`: it is a guest callable, and the whole point of
    // `emit_gui_property_set`'s binding work is that WHAT arrives here differs
    // per frontend — a bare function where `this` is ambient, an
    // `ecma:function.bind` result where the receiver had to be threaded. The
    // host calls it and does not care which; a narrower type would be a claim
    // this file cannot make.
    // `innerHTML`, both directions — DOM Parsing §2.3.
    //
    // The parser, the HTML grammar and the tree-builder all existed; every
    // entry point built a NEW document, so there was no way to say "make this
    // markup the contents of that element". This is the missing door, not a
    // missing engine.
    dom_fn(
        vm,
        "innerHtml",
        "inner-html",
        vec![doc(), node()],
        vec![ValType::String],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match crate::engine::apply(
                doc_arg(args, 0),
                crate::engine::DomOp::InnerHtml(node_arg(args, 1)),
            ) {
                crate::engine::DomValue::Text(html) => Value::String(Arc::from(html.as_str())),
                _ => Value::String(Arc::from("")),
            }
        }),
    );
    dom_fn(
        vm,
        "setInnerHtml",
        "set-inner-html",
        vec![doc(), node(), ValType::String],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            crate::engine::apply(
                doc_arg(args, 0),
                crate::engine::DomOp::SetInnerHtml {
                    node: node_arg(args, 1),
                    html: str_arg(args, 2),
                },
            );
            Value::Null
        }),
    );

    dom_fn(
        vm,
        "removeEventListener",
        "remove-event-listener",
        vec![doc(), node(), ValType::String, ValType::Any],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let cb = args.get(3).cloned().unwrap_or(Value::Undefined);
            remove_event_listener(doc_arg(args, 0), node_arg(args, 1), &str_arg(args, 2), &cb);
            Value::Null
        }),
    );

    dom_fn(
        vm,
        "addEventListener",
        "add-event-listener",
        vec![doc(), node(), ValType::String, ValType::Any],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let cb = args.get(3).cloned().unwrap_or(Value::Undefined);
            // **A null callback registers nothing** — DOM §2.7: "If callback is
            // null, then return." It is not an error and not a listener; a
            // browser accepts the call and does nothing, which is what lets a
            // framework hand over a whole slate of optional handlers and pass
            // `null` for the ones the program did not write.
            //
            // Registering it stored a listener that was null when the event
            // fired: every Flutter `ElevatedButton` declares `onPressed` AND
            // `onLongPress`, both wire to `click`, and a program that supplies
            // only the first made every click raise "null is not callable"
            // twice before reaching the handler that WAS there.
            if matches!(cb, Value::Null | Value::Undefined) {
                return Value::Null;
            }
            add_event_listener(doc_arg(args, 0), node_arg(args, 1), &str_arg(args, 2), cb);
            Value::Null
        }),
    );
}
