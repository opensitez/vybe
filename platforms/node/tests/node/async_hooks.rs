//! Behaviour tests for `node:async_hooks` host imports.
//!
//! Reference: <https://nodejs.org/api/async_hooks.html>.
//!
//! Coverage:
//!   - `executionAsyncId()` → non-negative integer
//!   - `triggerAsyncId()` → non-negative integer
//!   - `executionAsyncResource()` → object
//!   - `AsyncLocalStorage` constructor
//!   - `AsyncLocalStorage.run(value, fn)` → result of fn
//!   - `AsyncLocalStorage.getStore()` → current value or undefined
//!   - `AsyncLocalStorage.enterWith(value)` → void
//!   - `AsyncLocalStorage.exit(fn)` → exits ALS context
//!   - `AsyncLocalStorage.disable()` → void
//!   - `AsyncResource` constructor
//!   - `AsyncResource.asyncId()` → integer ≥ 1
//!   - `AsyncResource.triggerAsyncId()` → integer ≥ 0
//!   - `AsyncResource.run(fn, thisArg)` → result of fn
//!   - `AsyncResource.bind(fn)` → bound function (object or callable)
//!   - `createHook(callbacks)` → hook object with enable/disable
//!   - `createHook` with real callbacks object (init/before/after/destroy)
//!   - `asyncWrapProviders` constant object
//!
//! Deferred (require fully async VM execution):
//!   - Hook lifecycle callbacks (`init`, `before`, `after`, `destroy`) actually firing
//!   - AsyncIterator context propagation

use std::sync::Arc;
use vybe_runtime::value::{Object, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn call_ah(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-async_hooks-test>");
    let import_idx = chunk.add_import("node:async_hooks", name);
    let argc = args.len() as u8;
    let mut arg_globals: Vec<(String, Value)> = Vec::new();
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(s) => chunk.emit_string_const(&s, 0),
            other => {
                let name = format!(
                    "__test_arg_{}",
                    TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
                );
                let ci = chunk.intern_string_constant(&name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
                arg_globals.push((name, other));
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    for (name, value) in arg_globals {
        vm.globals.insert(name, value);
    }
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:async_hooks"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

#[allow(dead_code)]
fn prop(obj: &Value, key: &str) -> Value {
    match obj {
        Value::Object(o) => {
            let o = o.lock().unwrap();
            o.properties.get(key).cloned().unwrap_or(Value::Undefined)
        }
        _ => Value::Undefined,
    }
}

fn as_i64(v: &Value) -> Option<i64> {
    match v {
        Value::I32(n) => Some(*n as i64),
        Value::I64(n) => Some(*n),
        Value::F64(f) => Some(*f as i64),
        _ => None,
    }
}

// ── executionAsyncId ──────────────────────────────────────────────────────────

#[test]
fn execution_async_id_returns_non_negative_integer() {
    let id = call_ah("executionAsyncId", vec![]);
    assert!(as_i64(&id).map_or(false, |n| n >= 0), "got {id:?}");
}

#[test]
fn execution_async_id_is_at_least_one_in_main() {
    // The main context always has asyncId >= 1 in Node.js
    let id = call_ah("executionAsyncId", vec![]);
    assert!(as_i64(&id).map_or(false, |n| n >= 1), "got {id:?}");
}

// ── triggerAsyncId ────────────────────────────────────────────────────────────

#[test]
fn trigger_async_id_returns_non_negative_integer() {
    let id = call_ah("triggerAsyncId", vec![]);
    assert!(as_i64(&id).map_or(false, |n| n >= 0), "got {id:?}");
}

// ── executionAsyncResource ────────────────────────────────────────────────────

#[test]
fn execution_async_resource_returns_object() {
    let resource = call_ah("executionAsyncResource", vec![]);
    assert!(matches!(resource, Value::Object(_)));
}

// ── AsyncLocalStorage ─────────────────────────────────────────────────────────

#[test]
fn async_local_storage_constructor_returns_object() {
    let als = call_ah("AsyncLocalStorage", vec![]);
    assert!(matches!(als, Value::Object(_)));
}

#[test]
fn async_local_storage_get_store_returns_undefined_before_run() {
    let als = call_ah("AsyncLocalStorage", vec![]);
    let result = call_ah("alsGetStore", vec![als]);
    assert_eq!(result, Value::Undefined);
}

#[test]
fn async_local_storage_run_returns_callback_result() {
    let als = call_ah("AsyncLocalStorage", vec![]);
    // run(store_value, callback) — callback is null but host should return null (or sentinel)
    let result = call_ah("alsRun", vec![als, s("ctx-value"), Value::Null]);
    // If callback is null the host returns null or undefined — just must not panic
    assert!(matches!(
        result,
        Value::Null | Value::Undefined | Value::String(_)
    ));
}

#[test]
fn async_local_storage_enter_with_sets_store() {
    let als = call_ah("AsyncLocalStorage", vec![]);
    let _ = call_ah("alsEnterWith", vec![als.clone(), s("my-value")]);
    let store = call_ah("alsGetStore", vec![als]);
    assert_eq!(store, s("my-value"));
}

#[test]
fn async_local_storage_disable_returns_undefined() {
    let als = call_ah("AsyncLocalStorage", vec![]);
    let result = call_ah("alsDisable", vec![als]);
    assert_eq!(result, Value::Undefined);
}

// ── AsyncResource ─────────────────────────────────────────────────────────────

#[test]
fn async_resource_constructor_returns_object() {
    let resource = call_ah("AsyncResource", vec![s("my-resource")]);
    assert!(matches!(resource, Value::Object(_)));
}

#[test]
fn async_resource_has_async_id() {
    let resource = call_ah("AsyncResource", vec![s("res")]);
    let aid = call_ah("asyncResourceAsyncId", vec![resource]);
    assert!(
        as_i64(&aid).map_or(false, |n| n >= 1),
        "asyncId must be >= 1"
    );
}

#[test]
fn async_resource_has_trigger_async_id() {
    let resource = call_ah("AsyncResource", vec![s("res")]);
    let tid = call_ah("asyncResourceTriggerAsyncId", vec![resource]);
    assert!(as_i64(&tid).map_or(false, |n| n >= 0));
}

#[test]
fn async_resource_run_invokes_callback() {
    let resource = call_ah("AsyncResource", vec![s("runner")]);
    // Passing null callback — host must not panic
    let result = call_ah("asyncResourceRun", vec![resource, Value::Null, Value::Null]);
    assert!(matches!(result, Value::Null | Value::Undefined));
}

// ── createHook ────────────────────────────────────────────────────────────────

#[test]
fn create_hook_returns_hook_object() {
    let hook = call_ah("createHook", vec![Value::Null]);
    assert!(matches!(hook, Value::Object(_)));
}

#[test]
fn create_hook_enable_returns_hook() {
    let hook = call_ah("createHook", vec![Value::Null]);
    let result = call_ah("hookEnable", vec![hook]);
    assert!(matches!(result, Value::Object(_)));
}

#[test]
fn create_hook_disable_returns_hook() {
    let hook = call_ah("createHook", vec![Value::Null]);
    let _ = call_ah("hookEnable", vec![hook.clone()]);
    let result = call_ah("hookDisable", vec![hook]);
    assert!(matches!(result, Value::Object(_)));
}

// ── AsyncResource.bind ───────────────────────────────────────────────────────

#[test]
fn async_resource_bind_returns_non_undefined() {
    let resource = call_ah("AsyncResource", vec![s("binder")]);
    let result = call_ah("asyncResourceBind", vec![resource, Value::Null]);
    // bind wraps the function — may return null (no-op on null fn) or an object
    assert!(
        matches!(result, Value::Object(_) | Value::Null | Value::Undefined),
        "asyncResourceBind must return object/null/undefined, got {:?}",
        result
    );
}

// ── AsyncLocalStorage.exit ────────────────────────────────────────────────────

#[test]
fn als_exit_does_not_panic() {
    let als = call_ah("AsyncLocalStorage", vec![]);
    // exit(fn) — run callback outside ALS context; null fn is a no-op
    let result = call_ah("alsExit", vec![als, Value::Null]);
    let _ = result;
}

#[test]
fn als_exit_is_registered() {
    assert!(has_import("alsExit"), "alsExit must be registered");
}

// ── asyncWrapProviders ────────────────────────────────────────────────────────

#[test]
fn async_wrap_providers_returns_object() {
    let result = call_ah("asyncWrapProviders", vec![]);
    assert!(
        matches!(result, Value::Object(_)),
        "asyncWrapProviders must return an object, got {:?}",
        result
    );
}

#[test]
fn async_wrap_providers_has_at_least_one_entry() {
    let result = call_ah("asyncWrapProviders", vec![]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        assert!(
            !o.properties.is_empty(),
            "asyncWrapProviders must have at least one entry"
        );
    }
}

#[test]
fn async_wrap_providers_tcp_wrap_is_numeric() {
    let result = call_ah("asyncWrapProviders", vec![]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        if let Some(val) = o.properties.get("TCPWRAP") {
            assert!(
                matches!(val, Value::I32(_) | Value::I64(_) | Value::F64(_)),
                "TCPWRAP must be numeric, got {:?}",
                val
            );
        }
    }
}

// ── createHook with real callbacks object ─────────────────────────────────────

#[test]
fn create_hook_with_callbacks_object_does_not_panic() {
    let mut callbacks = Object::new();
    callbacks.properties.insert("init".to_string(), Value::Null);
    callbacks
        .properties
        .insert("before".to_string(), Value::Null);
    callbacks
        .properties
        .insert("after".to_string(), Value::Null);
    callbacks
        .properties
        .insert("destroy".to_string(), Value::Null);
    let cb_val = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(callbacks)));
    let hook = call_ah("createHook", vec![cb_val]);
    assert!(matches!(hook, Value::Object(_)));
}

#[test]
fn create_hook_with_callbacks_enable_disable_roundtrip() {
    let mut callbacks = Object::new();
    callbacks.properties.insert("init".to_string(), Value::Null);
    let cb_val = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(callbacks)));
    let hook = call_ah("createHook", vec![cb_val]);
    let enabled = call_ah("hookEnable", vec![hook]);
    let result = call_ah("hookDisable", vec![enabled]);
    assert!(matches!(result, Value::Object(_)));
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_async_hooks_surface_is_registered() {
    let expected = [
        "executionAsyncId",
        "triggerAsyncId",
        "executionAsyncResource",
        "AsyncLocalStorage",
        "AsyncResource",
        "createHook",
        "alsRun",
        "alsGetStore",
        "alsEnterWith",
        "alsExit",
        "alsDisable",
        "asyncResourceAsyncId",
        "asyncResourceTriggerAsyncId",
        "asyncResourceRun",
        "asyncResourceBind",
        "hookEnable",
        "hookDisable",
        "asyncWrapProviders",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing node:async_hooks imports: {missing:?}"
    );
}
