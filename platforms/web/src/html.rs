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
use vybe_runtime::{HostContext, VM, Value};

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

pub fn register(vm: &mut VM) {
    // ── Document ────────────────────────────────────────────────────────
    vm.register_host_fn(
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
    );
    vm.register_host_fn(
        "web:dom",
        "getElementById",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_node(apply(
                doc_arg(args, 0),
                DomOp::GetElementById(str_arg(args, 1)),
            ))
        }),
    );
    vm.register_host_fn(
        "web:dom",
        "getElementsByTagName",
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
    vm.register_host_fn(
        "web:html",
        "activeDocument",
        Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
            Value::F64(active_document() as f64)
        }),
    );

    // `document.body` — the document element every control hangs off.
    vm.register_host_fn(
        "web:html",
        "body",
        Box::new(move |_ctx: &mut HostContext, _args: &[Value]| element(DOCUMENT)),
    );
    vm.register_host_fn(
        "web:html",
        "title",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_text(apply(doc_arg(args, 0), DomOp::Title))
        }),
    );
    vm.register_host_fn(
        "web:html",
        "setTitle",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(doc_arg(args, 0), DomOp::SetTitle(str_arg(args, 1)));
            Value::Null
        }),
    );

    // ── Node ────────────────────────────────────────────────────────────
    vm.register_host_fn(
        "web:dom",
        "appendChild",
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
    vm.register_host_fn(
        "web:dom",
        "removeChild",
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
    vm.register_host_fn(
        "web:dom",
        "isConnected",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_bool(apply(
                doc_arg(args, 0),
                DomOp::IsConnected(node_arg(args, 1)),
            ))
        }),
    );
    vm.register_host_fn(
        "web:dom",
        "setTextContent",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::SetTextContent(node_arg(args, 1), str_arg(args, 2)),
            );
            Value::Null
        }),
    );
    vm.register_host_fn(
        "web:dom",
        "textContent",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_text(apply(
                doc_arg(args, 0),
                DomOp::TextContent(node_arg(args, 1)),
            ))
        }),
    );

    // ── Element: attributes ─────────────────────────────────────────────
    vm.register_host_fn(
        "web:dom",
        "setAttribute",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::SetAttribute(node_arg(args, 1), str_arg(args, 2), str_arg(args, 3)),
            );
            Value::Null
        }),
    );
    vm.register_host_fn(
        "web:dom",
        "getAttribute",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            // An absent attribute is `null`, per spec — not `""`.
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
    vm.register_host_fn(
        "web:dom",
        "toggleAttribute",
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
    vm.register_host_fn(
        "web:dom",
        "removeAttribute",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::RemoveAttribute(node_arg(args, 1), str_arg(args, 2)),
            );
            Value::Null
        }),
    );

    // ── CSSStyleDeclaration ─────────────────────────────────────────────
    vm.register_host_fn(
        "web:cssom",
        "setStyleProperty",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::SetStyleProperty(node_arg(args, 1), str_arg(args, 2), str_arg(args, 3)),
            );
            Value::Null
        }),
    );
    vm.register_host_fn(
        "web:cssom",
        "getStyleProperty",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_text(apply(
                doc_arg(args, 0),
                DomOp::GetStyleProperty(node_arg(args, 1), str_arg(args, 2)),
            ))
        }),
    );

    // ── HTMLInputElement / HTMLSelectElement IDL ────────────────────────
    vm.register_host_fn(
        "web:html",
        "setValue",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::SetValue(node_arg(args, 1), str_arg(args, 2)),
            );
            Value::Null
        }),
    );
    vm.register_host_fn(
        "web:html",
        "value",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_text(apply(doc_arg(args, 0), DomOp::Value(node_arg(args, 1))))
        }),
    );
    vm.register_host_fn(
        "web:html",
        "setChecked",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::SetChecked(node_arg(args, 1), truthy(args.get(2))),
            );
            Value::Null
        }),
    );
    vm.register_host_fn(
        "web:html",
        "checked",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_bool(apply(doc_arg(args, 0), DomOp::Checked(node_arg(args, 1))))
        }),
    );
    vm.register_host_fn(
        "web:html",
        "focus",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(doc_arg(args, 0), DomOp::Focus(node_arg(args, 1)));
            Value::Null
        }),
    );

    // `select.selectedIndex` — `-1` when nothing is selected, which is the
    // IDL's own answer and what every caller tests against.
    vm.register_host_fn(
        "web:html",
        "selectedIndex",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_number(apply(
                doc_arg(args, 0),
                DomOp::SelectedIndex(node_arg(args, 1)),
            ))
        }),
    );
    vm.register_host_fn(
        "web:html",
        "setSelectedIndex",
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
    vm.register_host_fn(
        "web:html",
        "itemText",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            as_text(apply(
                doc_arg(args, 0),
                DomOp::ItemText(node_arg(args, 1), num_arg(args, 2).max(0.0) as usize),
            ))
        }),
    );
    vm.register_host_fn(
        "web:html",
        "setItemText",
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
    vm.register_host_fn(
        "web:html",
        "addItem",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::AddItem(node_arg(args, 1), str_arg(args, 2)),
            );
            Value::Null
        }),
    );
    vm.register_host_fn(
        "web:html",
        "removeItem",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(
                doc_arg(args, 0),
                DomOp::RemoveItem(node_arg(args, 1), num_arg(args, 2) as usize),
            );
            Value::Null
        }),
    );
    vm.register_host_fn(
        "web:html",
        "clearItems",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            apply(doc_arg(args, 0), DomOp::ClearItems(node_arg(args, 1)));
            Value::Null
        }),
    );

    // ── EventTarget ─────────────────────────────────────────────────────
    vm.register_host_fn(
        "web:dom",
        "addEventListener",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let cb = args.get(3).cloned().unwrap_or(Value::Undefined);
            add_event_listener(doc_arg(args, 0), node_arg(args, 1), &str_arg(args, 2), cb);
            Value::Null
        }),
    );
}
