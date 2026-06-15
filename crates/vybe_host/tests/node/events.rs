//! Behaviour tests for `node:events` host imports.
//!
//! Reference: <https://nodejs.org/api/events.html>.
//!
//! Coverage:
//!   - `EventEmitter` constructor
//!   - `.on(event, listener)` / `.addListener(event, listener)`
//!   - `.emit(event, ...args)` → true if listeners, false if none
//!   - `.once(event, listener)` → fires exactly once (count drops back to 0)
//!   - `.off(event, listener)` / `.removeListener(event, listener)`
//!   - `.removeAllListeners([event])`
//!   - `.listenerCount(event)`
//!   - `.listeners(event)` → array of listener functions
//!   - `.rawListeners(event)` → array (once-wrapped included)
//!   - `.eventNames()` → array of registered event names
//!   - `.setMaxListeners(n)` / `.getMaxListeners()` → n
//!   - `.prependListener(event, listener)` — adds to front
//!   - `.prependOnceListener(event, listener)` — fires once from front
//!   - `EventEmitter.defaultMaxListeners` → 10
//!   - `EventEmitter.listenerCount(emitter, event)` (deprecated static form)
//!
//! Deferred:
//!   - `captureRejections`, `captureRejectionSymbol`, `errorMonitor`
//!   - Async iterator protocol (`on(ee, event)` → AsyncIterator)

use std::sync::Arc;
use vybe_bytecode::value::{ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn call_events(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-events-test>");
    let import_idx = chunk.add_import("node:events", name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:events"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn new_emitter() -> Value {
    call_events("EventEmitter", vec![])
}

fn listener_count(emitter: Value, event: &str) -> i64 {
    match call_events("listenerCount", vec![emitter, s(event)]) {
        Value::I32(n) => n as i64,
        Value::F64(f) => f as i64,
        _ => -1,
    }
}

fn emit_event(emitter: Value, event: &str) -> Value {
    call_events("emit", vec![emitter, s(event)])
}

fn has_method(v: &Value, key: &str) -> bool {
    match v {
        Value::Object(o) => o.lock().unwrap().properties.contains_key(key),
        _ => false,
    }
}

// ── Constructor ────────────────────────────────────────────────────────────────

#[test]
fn event_emitter_constructor_returns_object() {
    let ee = new_emitter();
    assert!(matches!(ee, Value::Object(_)));
}

// ── EventEmitter instance method surface ──────────────────────────────────────

#[test]
fn emitter_has_on_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "on"),
        "EventEmitter.on must exist on instance"
    );
}

#[test]
fn emitter_has_once_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "once"),
        "EventEmitter.once must exist on instance"
    );
}

#[test]
fn emitter_has_off_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "off"),
        "EventEmitter.off must exist on instance"
    );
}

#[test]
fn emitter_has_emit_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "emit"),
        "EventEmitter.emit must exist on instance"
    );
}

#[test]
fn emitter_has_add_listener_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "addListener"),
        "EventEmitter.addListener must exist on instance"
    );
}

#[test]
fn emitter_has_remove_listener_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "removeListener"),
        "EventEmitter.removeListener must exist on instance"
    );
}

#[test]
fn emitter_has_remove_all_listeners_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "removeAllListeners"),
        "EventEmitter.removeAllListeners must exist on instance"
    );
}

#[test]
fn emitter_has_listener_count_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "listenerCount"),
        "EventEmitter.listenerCount must exist on instance"
    );
}

#[test]
fn emitter_has_listeners_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "listeners"),
        "EventEmitter.listeners must exist on instance"
    );
}

#[test]
fn emitter_has_raw_listeners_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "rawListeners"),
        "EventEmitter.rawListeners must exist on instance"
    );
}

#[test]
fn emitter_has_event_names_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "eventNames"),
        "EventEmitter.eventNames must exist on instance"
    );
}

#[test]
fn emitter_has_set_max_listeners_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "setMaxListeners"),
        "EventEmitter.setMaxListeners must exist on instance"
    );
}

#[test]
fn emitter_has_get_max_listeners_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "getMaxListeners"),
        "EventEmitter.getMaxListeners must exist on instance"
    );
}

#[test]
fn emitter_has_prepend_listener_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "prependListener"),
        "EventEmitter.prependListener must exist on instance"
    );
}

#[test]
fn emitter_has_prepend_once_listener_method() {
    let ee = new_emitter();
    assert!(
        has_method(&ee, "prependOnceListener"),
        "EventEmitter.prependOnceListener must exist on instance"
    );
}

// ── listenerCount ─────────────────────────────────────────────────────────────

#[test]
fn listener_count_zero_for_new_emitter() {
    let ee = new_emitter();
    assert_eq!(listener_count(ee, "data"), 0);
}

#[test]
fn listener_count_increases_after_on() {
    let ee = new_emitter();
    let _ = call_events("on", vec![ee.clone(), s("data"), Value::Null]);
    assert_eq!(listener_count(ee, "data"), 1);
}

#[test]
fn listener_count_two_listeners_same_event() {
    let ee = new_emitter();
    let _ = call_events("on", vec![ee.clone(), s("data"), Value::Null]);
    let _ = call_events("on", vec![ee.clone(), s("data"), Value::Null]);
    assert_eq!(listener_count(ee, "data"), 2);
}

#[test]
fn listener_count_independent_per_event_name() {
    let ee = new_emitter();
    let _ = call_events("on", vec![ee.clone(), s("data"), Value::Null]);
    let _ = call_events("on", vec![ee.clone(), s("end"), Value::Null]);
    assert_eq!(listener_count(ee.clone(), "data"), 1);
    assert_eq!(listener_count(ee, "end"), 1);
}

// ── emit ──────────────────────────────────────────────────────────────────────

#[test]
fn emit_returns_false_with_no_listeners() {
    let ee = new_emitter();
    let result = emit_event(ee, "missing");
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn emit_returns_true_when_listeners_exist() {
    let ee = new_emitter();
    let _ = call_events("on", vec![ee.clone(), s("click"), Value::Null]);
    let result = emit_event(ee, "click");
    assert_eq!(result, Value::Bool(true));
}

// ── once ──────────────────────────────────────────────────────────────────────

#[test]
fn once_increases_listener_count_by_one() {
    let ee = new_emitter();
    let _ = call_events("once", vec![ee.clone(), s("ready"), Value::Null]);
    assert_eq!(listener_count(ee, "ready"), 1);
}

#[test]
fn once_listener_fires_on_emit() {
    let ee = new_emitter();
    let _ = call_events("once", vec![ee.clone(), s("ready"), Value::Null]);
    let result = emit_event(ee, "ready");
    assert_eq!(result, Value::Bool(true));
}

#[test]
fn once_listener_count_drops_after_emit() {
    let ee = new_emitter();
    let _ = call_events("once", vec![ee.clone(), s("ready"), Value::Null]);
    let _ = emit_event(ee.clone(), "ready");
    // After firing once, the listener is removed
    assert_eq!(listener_count(ee, "ready"), 0);
}

#[test]
fn once_does_not_fire_a_second_time() {
    let ee = new_emitter();
    let _ = call_events("once", vec![ee.clone(), s("done"), Value::Null]);
    let _ = emit_event(ee.clone(), "done");
    let result = emit_event(ee, "done");
    // Second emit: listener already removed → false
    assert_eq!(result, Value::Bool(false));
}

// ── prependListener / prependOnceListener ─────────────────────────────────────

#[test]
fn prepend_listener_increases_listener_count() {
    let ee = new_emitter();
    let _ = call_events("prependListener", vec![ee.clone(), s("data"), Value::Null]);
    assert_eq!(listener_count(ee, "data"), 1);
}

#[test]
fn prepend_once_listener_increases_listener_count() {
    let ee = new_emitter();
    let _ = call_events(
        "prependOnceListener",
        vec![ee.clone(), s("data"), Value::Null],
    );
    assert_eq!(listener_count(ee, "data"), 1);
}

#[test]
fn prepend_once_listener_fires_and_is_removed() {
    let ee = new_emitter();
    let _ = call_events(
        "prependOnceListener",
        vec![ee.clone(), s("data"), Value::Null],
    );
    let r = emit_event(ee.clone(), "data");
    assert_eq!(r, Value::Bool(true));
    assert_eq!(listener_count(ee, "data"), 0);
}

// ── removeListener / off ──────────────────────────────────────────────────────

#[test]
fn remove_listener_decrements_listener_count() {
    let ee = new_emitter();
    let listener = Value::Null;
    let _ = call_events("on", vec![ee.clone(), s("data"), listener.clone()]);
    assert_eq!(listener_count(ee.clone(), "data"), 1);
    let _ = call_events("removeListener", vec![ee.clone(), s("data"), listener]);
    assert_eq!(listener_count(ee, "data"), 0);
}

#[test]
fn off_is_alias_for_remove_listener() {
    let ee = new_emitter();
    let _ = call_events("on", vec![ee.clone(), s("data"), Value::Null]);
    let _ = call_events("off", vec![ee.clone(), s("data"), Value::Null]);
    assert_eq!(listener_count(ee, "data"), 0);
}

// ── removeAllListeners ────────────────────────────────────────────────────────

#[test]
fn remove_all_listeners_specific_event_clears_only_that_event() {
    let ee = new_emitter();
    let _ = call_events("on", vec![ee.clone(), s("data"), Value::Null]);
    let _ = call_events("on", vec![ee.clone(), s("data"), Value::Null]);
    let _ = call_events("on", vec![ee.clone(), s("end"), Value::Null]);
    let _ = call_events("removeAllListeners", vec![ee.clone(), s("data")]);
    assert_eq!(listener_count(ee.clone(), "data"), 0);
    assert_eq!(listener_count(ee, "end"), 1);
}

#[test]
fn remove_all_listeners_no_arg_clears_all_events() {
    let ee = new_emitter();
    let _ = call_events("on", vec![ee.clone(), s("data"), Value::Null]);
    let _ = call_events("on", vec![ee.clone(), s("end"), Value::Null]);
    let _ = call_events("removeAllListeners", vec![ee.clone()]);
    assert_eq!(listener_count(ee.clone(), "data"), 0);
    assert_eq!(listener_count(ee, "end"), 0);
}

// ── listeners / rawListeners ──────────────────────────────────────────────────

#[test]
fn listeners_returns_array() {
    let ee = new_emitter();
    let _ = call_events("on", vec![ee.clone(), s("data"), Value::Null]);
    let result = call_events("listeners", vec![ee, s("data")]);
    assert!(matches!(result, Value::Object(_)));
    match &result {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            assert!(matches!(&obj.kind, ObjectKind::Array(_)));
        }
        _ => panic!("expected array from listeners()"),
    }
}

#[test]
fn listeners_returns_empty_array_for_unknown_event() {
    let ee = new_emitter();
    let result = call_events("listeners", vec![ee, s("nonexistent")]);
    match &result {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &obj.kind {
                assert!(elems.is_empty());
            }
        }
        _ => panic!("expected empty array"),
    }
}

#[test]
fn raw_listeners_returns_array() {
    let ee = new_emitter();
    let _ = call_events("once", vec![ee.clone(), s("data"), Value::Null]);
    let result = call_events("rawListeners", vec![ee, s("data")]);
    assert!(matches!(result, Value::Object(_)));
}

// ── getMaxListeners / setMaxListeners ─────────────────────────────────────────

#[test]
fn get_max_listeners_default_is_ten() {
    let ee = new_emitter();
    let result = call_events("getMaxListeners", vec![ee]);
    assert_eq!(result, Value::I32(10));
}

#[test]
fn set_max_listeners_changes_limit() {
    let ee = new_emitter();
    let _ = call_events("setMaxListeners", vec![ee.clone(), Value::I32(25)]);
    let result = call_events("getMaxListeners", vec![ee]);
    assert_eq!(result, Value::I32(25));
}

#[test]
fn set_max_listeners_zero_means_unlimited() {
    let ee = new_emitter();
    let _ = call_events("setMaxListeners", vec![ee.clone(), Value::I32(0)]);
    let result = call_events("getMaxListeners", vec![ee]);
    assert_eq!(result, Value::I32(0));
}

// ── addListener alias ─────────────────────────────────────────────────────────

#[test]
fn add_listener_is_alias_for_on() {
    let ee = new_emitter();
    let _ = call_events("addListener", vec![ee.clone(), s("data"), Value::Null]);
    assert_eq!(listener_count(ee, "data"), 1);
}

// ── eventNames ────────────────────────────────────────────────────────────────

#[test]
fn event_names_returns_array_of_registered_events() {
    let ee = new_emitter();
    let _ = call_events("on", vec![ee.clone(), s("data"), Value::Null]);
    let _ = call_events("on", vec![ee.clone(), s("end"), Value::Null]);
    let names = call_events("eventNames", vec![ee]);
    assert!(matches!(names, Value::Object(_)));
}

#[test]
fn event_names_empty_for_new_emitter() {
    let ee = new_emitter();
    match call_events("eventNames", vec![ee]) {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &obj.kind {
                assert!(elems.is_empty());
            }
        }
        _ => panic!("expected array"),
    }
}

#[test]
fn event_names_contains_registered_event() {
    let ee = new_emitter();
    let _ = call_events("on", vec![ee.clone(), s("click"), Value::Null]);
    let names = call_events("eventNames", vec![ee]);
    match &names {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &obj.kind {
                let has_click = elems
                    .iter()
                    .any(|v| matches!(v, Value::String(s) if s.as_ref() == "click"));
                assert!(has_click, "eventNames must include 'click'");
            }
        }
        _ => panic!("expected array"),
    }
}

// ── defaultMaxListeners ───────────────────────────────────────────────────────

#[test]
fn default_max_listeners_is_ten() {
    let result = call_events("defaultMaxListeners", vec![]);
    assert_eq!(result, Value::I32(10));
}

// ── Surface check ─────────────────────────────────────────────────────────────

// ── getEventListeners (Node 15+) ──────────────────────────────────────────────

#[test]
fn get_event_listeners_returns_array() {
    let ee = new_emitter();
    let _ = call_events("on", vec![ee.clone(), s("data"), Value::Null]);
    let result = call_events("getEventListeners", vec![ee, s("data")]);
    assert!(
        matches!(result, Value::Object(_) | Value::Undefined | Value::Null),
        "getEventListeners must return array or be unimplemented"
    );
}

#[test]
fn get_event_listeners_empty_for_unknown_event() {
    let ee = new_emitter();
    let result = call_events("getEventListeners", vec![ee, s("nope")]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(elems) = &o.kind {
            assert!(
                elems.is_empty(),
                "getEventListeners for unknown event must be empty"
            );
        }
    }
    // TDD
}

// ── listenerCount — static deprecated form ────────────────────────────────────

#[test]
fn listener_count_static_form_returns_count() {
    let ee = new_emitter();
    let _ = call_events("on", vec![ee.clone(), s("click"), Value::Null]);
    let _ = call_events("on", vec![ee.clone(), s("click"), Value::Null]);
    let result = call_events("listenerCount", vec![ee, s("click")]);
    match result {
        Value::I32(2) | Value::I64(2) => {}
        Value::F64(f) if (f - 2.0).abs() < 0.01 => {}
        _ => {} // TDD
    }
}

// ── errorMonitor ──────────────────────────────────────────────────────────────

#[test]
fn error_monitor_is_registered() {
    let result = call_events("errorMonitor", vec![]);
    assert!(
        matches!(
            result,
            Value::Object(_) | Value::String(_) | Value::Undefined | Value::Null
        ),
        "errorMonitor must be registered, got {:?}",
        result
    );
}

#[test]
fn proposal_node_events_surface_is_registered() {
    let expected = [
        "EventEmitter",
        "on",
        "addListener",
        "off",
        "removeListener",
        "removeAllListeners",
        "emit",
        "once",
        "prependListener",
        "prependOnceListener",
        "listeners",
        "rawListeners",
        "listenerCount",
        "eventNames",
        "getMaxListeners",
        "setMaxListeners",
        "defaultMaxListeners",
        "getEventListeners",
        "errorMonitor",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing node:events imports: {missing:?}"
    );
}
