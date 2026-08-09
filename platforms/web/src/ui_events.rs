//! W3C UI Events — the `web:ui-events` host functions.
//!
//! Exposure only. The queue and the pointer state it derives live in the
//! engine (see [`engine`](crate::engine)); what this file owns is the
//! marshalling a toolkit must not know about — turning an event into a guest
//! object and reading one back.
//!
//! Event objects carry the spec's own field names and value conventions —
//! `type`, `key`, `code`, `clientX`, `clientY`, `button`, `buttons`,
//! `deltaY`, `ctrlKey`/`shiftKey`/`altKey`/`metaKey` — so a guest that knows
//! the DOM needs no translation. Adapters that speak another vocabulary (the
//! SDL bridge maps `mousedown` → `SDL_MOUSEBUTTONDOWN`, DOM's 0-based
//! `button` → SDL's 1-based) do that translation on THEIR side, which is what
//! keeps this one standard.
//!
//! `dispatchEvent` is `EventTarget.dispatchEvent`: it injects a synthetic
//! event, exactly as the DOM allows, which also makes the whole pipeline
//! testable with no window open. `pollEvent` is the one non-DOM primitive —
//! the drain a polling guest (C, SDL, a game loop) needs where a browser
//! would invoke listeners.

use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

use crate::engine::{EventOp, EventValue, UiEventFields, events};

fn obj(props: Vec<(&str, Value)>) -> Value {
    let mut o = Object::new();
    for (k, v) in props {
        o.properties.insert(k.into(), v);
    }
    Value::Object(vybe_runtime::heap::alloc(o))
}

/// Build the guest-visible event object — W3C attribute names.
fn event_object(e: &UiEventFields) -> Value {
    obj(vec![
        ("type", Value::String(e.kind.as_str().into())),
        ("key", Value::String(e.key.as_str().into())),
        ("code", Value::String(e.code.as_str().into())),
        ("keyCode", Value::I32(e.key_code)),
        ("which", Value::I32(e.key_code)),
        ("clientX", Value::I32(e.client_x)),
        ("clientY", Value::I32(e.client_y)),
        ("button", Value::I32(e.button)),
        ("buttons", Value::I32(e.buttons)),
        ("deltaY", Value::F64(e.delta_y)),
        ("ctrlKey", Value::Bool(e.ctrl_key)),
        ("shiftKey", Value::Bool(e.shift_key)),
        ("altKey", Value::Bool(e.alt_key)),
        ("metaKey", Value::Bool(e.meta_key)),
        ("repeat", Value::Bool(e.repeat)),
    ])
}

fn get_str(o: &Object, k: &str) -> String {
    o.properties
        .get(k)
        .map(|v| format!("{}", v))
        .filter(|s| s != "null" && s != "undefined")
        .unwrap_or_default()
}

fn get_i32(o: &Object, k: &str) -> i32 {
    o.properties.get(k).map(|v| v.as_i32()).unwrap_or(0)
}

/// DOM boolean attributes with JS TRUTHINESS, not strict `Value::Bool`.
///
/// `Value::as_bool()` is `matches!(v, Bool(true))` — an `I32(1)` reads as
/// false. A guest that computes `ctrlKey` with a bitmask (which is what any
/// C/SDL adapter does) stores a number, and the modifier silently vanished.
/// The web platform's own coercion is truthiness, so use it here.
fn get_bool(o: &Object, k: &str) -> bool {
    match o.properties.get(k) {
        Some(Value::Bool(b)) => *b,
        Some(Value::I32(n)) => *n != 0,
        Some(Value::I64(n)) => *n != 0,
        Some(Value::F64(f)) => *f != 0.0 && !f.is_nan(),
        Some(Value::String(s)) => !s.is_empty(),
        Some(Value::Null) | Some(Value::Undefined) | None => false,
        Some(_) => true,
    }
}

/// Read a guest event object back into spec fields — the inverse of
/// [`event_object`], used by `dispatchEvent`.
fn event_from_value(v: Option<&Value>) -> Option<UiEventFields> {
    let Some(Value::Object(o)) = v else {
        return None;
    };
    let o = o.lock().unwrap();
    Some(UiEventFields {
        kind: get_str(&o, "type"),
        key: get_str(&o, "key"),
        code: get_str(&o, "code"),
        key_code: get_i32(&o, "keyCode"),
        client_x: get_i32(&o, "clientX"),
        client_y: get_i32(&o, "clientY"),
        button: get_i32(&o, "button"),
        buttons: get_i32(&o, "buttons"),
        delta_y: o
            .properties
            .get("deltaY")
            .map(|v| v.as_f64())
            .unwrap_or(0.0),
        ctrl_key: get_bool(&o, "ctrlKey"),
        shift_key: get_bool(&o, "shiftKey"),
        alt_key: get_bool(&o, "altKey"),
        meta_key: get_bool(&o, "metaKey"),
        repeat: get_bool(&o, "repeat"),
    })
}

fn tracing() -> bool {
    std::env::var_os("VYBE_TRACE_INPUT").is_some()
}

pub fn register(vm: &mut VM) {
    // new KeyboardEvent(type) / new MouseEvent(type) — the DOM constructs
    // events before dispatching them; this returns a zeroed event object of
    // the given type for the caller to fill and hand to `dispatchEvent`.
    vm.register_host_fn(
        "web:ui-events",
        "newEvent",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let kind = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            event_object(&UiEventFields {
                kind,
                ..UiEventFields::default()
            })
        }),
    );

    // EventTarget.dispatchEvent(event) → bool (true = not cancelled).
    vm.register_host_fn(
        "web:ui-events",
        "dispatchEvent",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match event_from_value(args.first()) {
                Some(evt) => {
                    if tracing() {
                        eprintln!(
                            "[dispatch] type={:?} keyCode={} shift={} ctrl={} rawmod={}",
                            evt.kind, evt.key_code, evt.shift_key, evt.ctrl_key, evt.delta_y
                        );
                    }
                    events(EventOp::Dispatch(evt));
                    Value::Bool(true)
                }
                None => Value::Bool(false),
            }
        }),
    );

    // pollEvent() → the oldest queued event, or null. NOT a DOM method: the
    // drain a polling guest needs where a browser invokes listeners.
    vm.register_host_fn(
        "web:ui-events",
        "pollEvent",
        Box::new(
            move |_ctx: &mut HostContext, _args: &[Value]| match events(EventOp::Poll) {
                EventValue::Event(e) => {
                    if tracing() {
                        eprintln!(
                            "[poll] type={:?} keyCode={} x={} y={}",
                            e.kind, e.key_code, e.client_x, e.client_y
                        );
                    }
                    event_object(&e)
                }
                _ => Value::Null,
            },
        ),
    );

    // pendingEvents() → queue depth (lets a loop drain without polling blind).
    vm.register_host_fn(
        "web:ui-events",
        "pendingEvents",
        Box::new(
            move |_ctx: &mut HostContext, _args: &[Value]| match events(EventOp::Pending) {
                EventValue::Count(n) => Value::I32(n as i32),
                _ => Value::I32(0),
            },
        ),
    );

    // pointerState() → {clientX, clientY, buttons, ctrlKey, …} — what a page
    // tracks from the event stream, for guests that sample instead of listen.
    vm.register_host_fn(
        "web:ui-events",
        "pointerState",
        Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
            match events(EventOp::PointerState) {
                EventValue::Pointer {
                    client_x,
                    client_y,
                    buttons,
                    ctrl_key,
                    shift_key,
                    alt_key,
                    meta_key,
                } => obj(vec![
                    ("clientX", Value::I32(client_x)),
                    ("clientY", Value::I32(client_y)),
                    ("buttons", Value::I32(buttons)),
                    ("ctrlKey", Value::Bool(ctrl_key)),
                    ("shiftKey", Value::Bool(shift_key)),
                    ("altKey", Value::Bool(alt_key)),
                    ("metaKey", Value::Bool(meta_key)),
                ]),
                _ => obj(vec![
                    ("clientX", Value::I32(0)),
                    ("clientY", Value::I32(0)),
                    ("buttons", Value::I32(0)),
                    ("ctrlKey", Value::Bool(false)),
                    ("shiftKey", Value::Bool(false)),
                    ("altKey", Value::Bool(false)),
                    ("metaKey", Value::Bool(false)),
                ]),
            }
        }),
    );
}
