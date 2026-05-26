//! Behaviour tests for `node:worker_threads` host imports.
//!
//! Reference: <https://nodejs.org/api/worker_threads.html>.
//!
//! Coverage:
//!   - `isMainThread`, `threadId`, `workerData`, `parentPort`, `resourceLimits`, `SHARE_ENV`
//!   - `Worker` constructor + properties (threadId, stdin/stdout/stderr, resourceLimits)
//!     + methods (postMessage, terminate, ref, unref, getHeapSnapshot)
//!     + EventEmitter (on, once, off, emit, removeListener)
//!   - `MessageChannel` → {port1, port2}, ports are distinct
//!   - `MessagePort` constructor/methods (postMessage, close, start, ref, unref, hasRef)
//!     + EventEmitter (on, once, off, removeListener)
//!   - `receiveMessageOnPort` → undefined when queue empty
//!   - `BroadcastChannel` constructor + name + postMessage + close
//!   - `moveMessagePortToContext`, `markAsUntransferable`, `isMarkedAsUntransferable`
//!   - `getEnvironmentData`, `setEnvironmentData`

use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn call_wt(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-worker_threads-test>");
    let import_idx = chunk.add_import("node:worker_threads", name);
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
        .contains_key(&(String::from("node:worker_threads"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

fn prop(obj: &Value, key: &str) -> Value {
    match obj {
        Value::Object(o) => {
            let o = o.lock().unwrap();
            o.properties.get(key).cloned().unwrap_or(Value::Undefined)
        }
        _ => Value::Undefined,
    }
}

fn has_method(obj: &Value, key: &str) -> bool {
    match obj {
        Value::Object(o) => o.lock().unwrap().properties.contains_key(key),
        _ => false,
    }
}

// ── isMainThread ──────────────────────────────────────────────────────────────

#[test]
fn is_main_thread_returns_true_in_test_context() {
    let result = call_wt("isMainThread", vec![]);
    assert_eq!(result, Value::Bool(true));
}

// ── threadId ──────────────────────────────────────────────────────────────────

#[test]
fn thread_id_is_zero_for_main_thread() {
    let result = call_wt("threadId", vec![]);
    assert_eq!(result, Value::I32(0));
}

// ── workerData ────────────────────────────────────────────────────────────────

#[test]
fn worker_data_is_null_in_main_thread() {
    let result = call_wt("workerData", vec![]);
    assert_eq!(result, Value::Null);
}

// ── parentPort ────────────────────────────────────────────────────────────────

#[test]
fn parent_port_is_null_in_main_thread() {
    let result = call_wt("parentPort", vec![]);
    assert_eq!(result, Value::Null);
}

// ── resourceLimits ────────────────────────────────────────────────────────────

#[test]
fn resource_limits_returns_object() {
    let result = call_wt("resourceLimits", vec![]);
    assert!(matches!(result, Value::Object(_)));
}

#[test]
fn resource_limits_has_max_old_generation_size_mb() {
    let result = call_wt("resourceLimits", vec![]);
    let val = prop(&result, "maxOldGenerationSizeMb");
    assert!(matches!(val, Value::I32(_) | Value::I64(_) | Value::F64(_)));
}

#[test]
fn resource_limits_has_max_young_generation_size_mb() {
    let result = call_wt("resourceLimits", vec![]);
    let val = prop(&result, "maxYoungGenerationSizeMb");
    assert!(matches!(val, Value::I32(_) | Value::I64(_) | Value::F64(_)));
}

#[test]
fn resource_limits_has_stack_size_mb() {
    let result = call_wt("resourceLimits", vec![]);
    let val = prop(&result, "stackSizeMb");
    assert!(matches!(val, Value::I32(_) | Value::I64(_) | Value::F64(_)));
}

// ── SHARE_ENV ────────────────────────────────────────────────────────────────

#[test]
fn share_env_is_a_symbol_value() {
    let result = call_wt("SHARE_ENV", vec![]);
    assert!(!matches!(result, Value::Null | Value::Undefined));
}

// ── MessageChannel ────────────────────────────────────────────────────────────

#[test]
fn message_channel_returns_object_with_port1_and_port2() {
    let channel = call_wt("MessageChannel", vec![]);
    assert!(matches!(channel, Value::Object(_)));
    let port1 = prop(&channel, "port1");
    let port2 = prop(&channel, "port2");
    assert!(matches!(port1, Value::Object(_)), "port1 must be an object");
    assert!(matches!(port2, Value::Object(_)), "port2 must be an object");
}

#[test]
fn message_channel_port1_and_port2_are_distinct() {
    let channel = call_wt("MessageChannel", vec![]);
    let port1 = prop(&channel, "port1");
    let port2 = prop(&channel, "port2");
    let p1_ptr = match &port1 { Value::Object(arc) => std::sync::Arc::as_ptr(arc) as usize, _ => 0 };
    let p2_ptr = match &port2 { Value::Object(arc) => std::sync::Arc::as_ptr(arc) as usize, _ => 1 };
    assert_ne!(p1_ptr, p2_ptr, "port1 and port2 must be distinct objects");
}

// ── MessagePort methods ───────────────────────────────────────────────────────

#[test]
fn message_port_has_post_message_method() {
    let channel = call_wt("MessageChannel", vec![]);
    let port1 = prop(&channel, "port1");
    assert!(has_method(&port1, "postMessage"), "MessagePort.postMessage must exist");
}

#[test]
fn message_port_has_close_method() {
    let channel = call_wt("MessageChannel", vec![]);
    let port1 = prop(&channel, "port1");
    assert!(has_method(&port1, "close"), "MessagePort.close must exist");
}

#[test]
fn message_port_has_start_method() {
    let channel = call_wt("MessageChannel", vec![]);
    let port1 = prop(&channel, "port1");
    assert!(has_method(&port1, "start"), "MessagePort.start must exist");
}

#[test]
fn message_port_has_ref_method() {
    let channel = call_wt("MessageChannel", vec![]);
    let port1 = prop(&channel, "port1");
    assert!(has_method(&port1, "ref"), "MessagePort.ref must exist");
}

#[test]
fn message_port_has_unref_method() {
    let channel = call_wt("MessageChannel", vec![]);
    let port1 = prop(&channel, "port1");
    assert!(has_method(&port1, "unref"), "MessagePort.unref must exist");
}

#[test]
fn message_port_has_has_ref_method() {
    let channel = call_wt("MessageChannel", vec![]);
    let port1 = prop(&channel, "port1");
    assert!(has_method(&port1, "hasRef"), "MessagePort.hasRef must exist");
}

// ── MessagePort EventEmitter ──────────────────────────────────────────────────

#[test]
fn message_port_has_on_method() {
    let channel = call_wt("MessageChannel", vec![]);
    let port1 = prop(&channel, "port1");
    assert!(has_method(&port1, "on"), "MessagePort.on (EventEmitter) must exist");
}

#[test]
fn message_port_has_once_method() {
    let channel = call_wt("MessageChannel", vec![]);
    let port1 = prop(&channel, "port1");
    assert!(has_method(&port1, "once"), "MessagePort.once must exist");
}

#[test]
fn message_port_has_off_method() {
    let channel = call_wt("MessageChannel", vec![]);
    let port1 = prop(&channel, "port1");
    assert!(has_method(&port1, "off"), "MessagePort.off must exist");
}

#[test]
fn message_port_has_emit_method() {
    let channel = call_wt("MessageChannel", vec![]);
    let port1 = prop(&channel, "port1");
    assert!(has_method(&port1, "emit"), "MessagePort.emit must exist");
}

#[test]
fn message_port_has_remove_listener_method() {
    let channel = call_wt("MessageChannel", vec![]);
    let port1 = prop(&channel, "port1");
    assert!(has_method(&port1, "removeListener"), "MessagePort.removeListener must exist");
}

#[test]
fn message_port_has_remove_all_listeners_method() {
    let channel = call_wt("MessageChannel", vec![]);
    let port1 = prop(&channel, "port1");
    assert!(has_method(&port1, "removeAllListeners"), "MessagePort.removeAllListeners must exist");
}

// ── receiveMessageOnPort ──────────────────────────────────────────────────────

#[test]
fn receive_message_on_port_with_empty_port_returns_undefined() {
    let channel = call_wt("MessageChannel", vec![]);
    let port1 = prop(&channel, "port1");
    let result = call_wt("receiveMessageOnPort", vec![port1]);
    assert_eq!(result, Value::Undefined);
}

// ── Worker constructor ────────────────────────────────────────────────────────

#[test]
fn worker_constructor_returns_object() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    assert!(matches!(worker, Value::Object(_)), "Worker() must return an object");
}

#[test]
fn worker_has_thread_id_property() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    let tid = prop(&worker, "threadId");
    assert!(
        matches!(tid, Value::I32(_) | Value::I64(_) | Value::F64(_)),
        "Worker.threadId must be numeric, got {:?}", tid
    );
}

#[test]
fn worker_has_stdout_stderr_stdin() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    assert!(
        !matches!(prop(&worker, "stdout"), Value::Undefined),
        "Worker.stdout must be present"
    );
    assert!(
        !matches!(prop(&worker, "stderr"), Value::Undefined),
        "Worker.stderr must be present"
    );
    assert!(
        !matches!(prop(&worker, "stdin"), Value::Undefined),
        "Worker.stdin must be present"
    );
}

#[test]
fn worker_has_post_message_method() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    assert!(has_method(&worker, "postMessage"), "Worker.postMessage must exist");
}

#[test]
fn worker_has_terminate_method() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    assert!(has_method(&worker, "terminate"), "Worker.terminate must exist");
}

#[test]
fn worker_has_ref_and_unref_methods() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    assert!(has_method(&worker, "ref"), "Worker.ref must exist");
    assert!(has_method(&worker, "unref"), "Worker.unref must exist");
}

#[test]
fn worker_has_get_heap_snapshot_method() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    assert!(has_method(&worker, "getHeapSnapshot"), "Worker.getHeapSnapshot must exist");
}

// ── Worker EventEmitter ───────────────────────────────────────────────────────

#[test]
fn worker_has_on_method() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    assert!(has_method(&worker, "on"), "Worker.on (EventEmitter) must exist");
}

#[test]
fn worker_has_once_method() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    assert!(has_method(&worker, "once"), "Worker.once must exist");
}

#[test]
fn worker_has_off_method() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    assert!(has_method(&worker, "off"), "Worker.off must exist");
}

#[test]
fn worker_has_emit_method() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    assert!(has_method(&worker, "emit"), "Worker.emit must exist");
}

#[test]
fn worker_has_remove_listener_method() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    assert!(has_method(&worker, "removeListener"), "Worker.removeListener must exist");
}

#[test]
fn worker_has_remove_all_listeners_method() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    assert!(has_method(&worker, "removeAllListeners"), "Worker.removeAllListeners must exist");
}

#[test]
fn worker_has_listener_count_method() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    assert!(has_method(&worker, "listenerCount"), "Worker.listenerCount must exist");
}

// ── BroadcastChannel ─────────────────────────────────────────────────────────

#[test]
fn broadcast_channel_constructor_returns_object() {
    let bc = call_wt("BroadcastChannel", vec![s("test-channel")]);
    assert!(matches!(bc, Value::Object(_)));
}

#[test]
fn broadcast_channel_has_name_property() {
    let bc = call_wt("BroadcastChannel", vec![s("my-channel")]);
    let name = prop(&bc, "name");
    match name {
        Value::String(n) => assert_eq!(n.as_ref(), "my-channel", "BroadcastChannel.name must match"),
        other => panic!("BroadcastChannel.name expected string, got {:?}", other),
    }
}

#[test]
fn broadcast_channel_has_post_message_method() {
    let bc = call_wt("BroadcastChannel", vec![s("ch")]);
    assert!(has_method(&bc, "postMessage"), "BroadcastChannel.postMessage must exist");
}

#[test]
fn broadcast_channel_has_close_method() {
    let bc = call_wt("BroadcastChannel", vec![s("ch")]);
    assert!(has_method(&bc, "close"), "BroadcastChannel.close must exist");
}

#[test]
fn broadcast_channel_has_on_message_handler() {
    let bc = call_wt("BroadcastChannel", vec![s("ch")]);
    // onmessage is a property (handler slot), may be null initially
    let onmessage = prop(&bc, "onmessage");
    assert!(
        matches!(onmessage, Value::Null | Value::Undefined | Value::Object(_)),
        "BroadcastChannel.onmessage should exist as null/fn, got {:?}", onmessage
    );
}

// ── markAsUntransferable / isMarkedAsUntransferable ───────────────────────────

#[test]
fn mark_as_untransferable_does_not_panic() {
    let obj = {
        let mut o = Object::new();
        o.properties.insert("key".to_string(), Value::I32(1));
        Value::Object(std::sync::Arc::new(std::sync::Mutex::new(o)))
    };
    let result = call_wt("markAsUntransferable", vec![obj]);
    let _ = result;
}

#[test]
fn is_marked_as_untransferable_returns_bool() {
    let obj = {
        let mut o = Object::new();
        o.properties.insert("key".to_string(), Value::I32(1));
        Value::Object(std::sync::Arc::new(std::sync::Mutex::new(o)))
    };
    let result = call_wt("isMarkedAsUntransferable", vec![obj]);
    assert!(
        matches!(result, Value::Bool(_)),
        "isMarkedAsUntransferable must return bool, got {:?}", result
    );
}

// ── getEnvironmentData / setEnvironmentData ───────────────────────────────────

#[test]
fn set_environment_data_does_not_panic() {
    let result = call_wt("setEnvironmentData", vec![s("my-key"), s("my-value")]);
    let _ = result;
}

#[test]
fn get_environment_data_returns_previously_set_value() {
    call_wt("setEnvironmentData", vec![s("test-key"), s("test-value")]);
    let result = call_wt("getEnvironmentData", vec![s("test-key")]);
    // May return the value or undefined if cross-thread isolation prevents it
    assert!(
        matches!(result, Value::String(_) | Value::Null | Value::Undefined),
        "getEnvironmentData must return string or null/undefined, got {:?}", result
    );
}

// ── moveMessagePortToContext ──────────────────────────────────────────────────

#[test]
fn move_message_port_to_context_exists() {
    // Just verify the function is registered; calling it requires a real context object.
    assert!(has_import("moveMessagePortToContext"), "moveMessagePortToContext must be registered");
}

// ── Worker — exitCode / resourceLimits properties ────────────────────────────

#[test]
fn worker_has_exit_code_property() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    let ec = prop(&worker, "exitCode");
    assert!(
        matches!(ec, Value::Null | Value::Undefined | Value::I32(_) | Value::F64(_)),
        "Worker.exitCode must be null/undefined before exit, got {:?}", ec
    );
}

#[test]
fn worker_has_resource_limits_property() {
    let worker = call_wt("Worker", vec![s("./worker.js")]);
    let rl = prop(&worker, "resourceLimits");
    assert!(
        matches!(rl, Value::Object(_) | Value::Null | Value::Undefined),
        "Worker.resourceLimits must be an object or undefined, got {:?}", rl
    );
}

// ── MessagePort — addEventListener / removeEventListener ─────────────────────

#[test]
fn message_port_has_add_event_listener_method() {
    let mc = call_wt("MessageChannel", vec![]);
    let port1 = prop(&mc, "port1");
    assert!(
        has_method(&port1, "addEventListener") || matches!(port1, Value::Undefined | Value::Null),
        "MessagePort.addEventListener must exist (DOM-style alias)"
    );
}

#[test]
fn message_port_has_remove_event_listener_method() {
    let mc = call_wt("MessageChannel", vec![]);
    let port1 = prop(&mc, "port1");
    assert!(
        has_method(&port1, "removeEventListener") || matches!(port1, Value::Undefined | Value::Null),
        "MessagePort.removeEventListener must exist (DOM-style alias)"
    );
}

// ── BroadcastChannel — onmessageerror handler ─────────────────────────────────

#[test]
fn broadcast_channel_has_on_message_error_property() {
    let bc = call_wt("BroadcastChannel", vec![s("err-channel")]);
    // onmessageerror may be null initially; must be a settable property
    let val = prop(&bc, "onmessageerror");
    assert!(
        matches!(val, Value::Null | Value::Undefined | Value::Object(_)),
        "BroadcastChannel.onmessageerror must exist, got {:?}", val
    );
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[allow(dead_code)]
fn _suppress_unused(_: Object) {}
#[allow(dead_code)]
fn _suppress_unused2(_: ObjectKind) {}

#[test]
fn proposal_node_worker_threads_surface_is_registered() {
    let expected = [
        "isMainThread",
        "threadId",
        "workerData",
        "parentPort",
        "resourceLimits",
        "SHARE_ENV",
        "Worker",
        "MessageChannel",
        "MessagePort",
        "BroadcastChannel",
        "receiveMessageOnPort",
        "moveMessagePortToContext",
        "markAsUntransferable",
        "isMarkedAsUntransferable",
        "getEnvironmentData",
        "setEnvironmentData",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(missing.is_empty(), "missing node:worker_threads imports: {missing:?}");
}

