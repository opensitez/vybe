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
use std::sync::Mutex;

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
    listeners()
        .lock()
        .unwrap()
        .entry((document, node, kind.to_ascii_lowercase()))
        .or_default()
        .push(callback);
}

pub fn listeners_for(document: DocumentId, node: NodeId, kind: &str) -> Vec<Value> {
    listeners()
        .lock()
        .unwrap()
        .get(&(document, node, kind.to_ascii_lowercase()))
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

// ── argument decoding ───────────────────────────────────────────────────

fn doc_arg_id(args: &[Value], idx: usize) -> Option<DocumentId> {
    args.get(idx).map(|v| v.as_f64() as DocumentId)
}

fn doc_arg(args: &[Value], idx: usize) -> DocumentId {
    args.get(idx).map(|v| v.as_f64() as DocumentId).unwrap_or(0)
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

/// Wrap a node as the element object guest code holds.
fn element(node: NodeId) -> Value {
    let mut o = Object::new();
    o.properties
        .insert("__node".into(), Value::F64(node as f64));
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

fn as_node(v: DomValue) -> Value {
    match v {
        DomValue::Node(n) => element(n),
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
fn dom_fn(
    vm: &mut VM,
    name: &str,
    kebab: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(
        HostFnDecl::new("web:dom", name, call)
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
        HostFnDecl::new("web:html", name, call)
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
        HostFnDecl::new("web:cssom", name, call)
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
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                as_node(apply(
                    doc_arg(args, 0),
                    DomOp::CreateElement {
                        tag: str_arg(args, 1),
                        input_type: str_arg(args, 2),
                    },
                ))
            }),
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
    vm.register_host(
        HostFnDecl::new(
            "web:dom",
            "getElementById",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                as_node(apply(
                    doc_arg(args, 0),
                    DomOp::GetElementById(str_arg(args, 1)),
                ))
            }),
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
                    let items: Vec<Value> = ns.into_iter().map(element).collect();
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
                Value::F64(active_document() as f64)
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
            borrows_self: true,
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
        Box::new(move |_ctx: &mut HostContext, _args: &[Value]| element(DOCUMENT)),
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
    dom_fn(
        vm,
        "addEventListener",
        "add-event-listener",
        vec![doc(), node(), ValType::String, ValType::Any],
        vec![],
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let cb = args.get(3).cloned().unwrap_or(Value::Undefined);
            add_event_listener(doc_arg(args, 0), node_arg(args, 1), &str_arg(args, 2), cb);
            Value::Null
        }),
    );
}
