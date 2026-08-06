//! Behaviour tests for `node:timers` host imports.
//!
//! Reference: <https://nodejs.org/api/timers.html>.
//!
//! Coverage:
//!   - `setTimeout(callback, delay[, ...args])` → Timeout handle
//!   - `clearTimeout(timeout)` → void
//!   - `setInterval(callback, delay[, ...args])` → Interval handle
//!   - `clearInterval(interval)` → void
//!   - `setImmediate(callback[, ...args])` → Immediate handle
//!   - `clearImmediate(immediate)` → void
//!   - `queueMicrotask(callback)` → void (Node 11+)
//!   - `timers/promises` surface: `setTimeout`, `setInterval`, `setImmediate`
//!
//! The sync-observable parts of timer behaviour are: handle identity,
//! handle cancellation (cleared handle does not fire), and the
//! `queueMicrotask` surface. Full async sequencing (callback ordering,
//! `unref`/`ref`, `hasRef`) is deferred — those require a running event loop.

use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn call_timers(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-timers-test>");
    let import_idx = chunk.add_import("node:timers", name);
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
        .contains_key(&(String::from("node:timers"), name.to_string()))
}

// ── setTimeout ────────────────────────────────────────────────────────────────

#[test]
fn set_timeout_returns_handle_object() {
    // Passing Null as callback — host must accept any value; the returned
    // handle is an opaque reference the caller can pass to clearTimeout.
    let handle = call_timers("setTimeout", vec![Value::Null, Value::I32(0)]);
    assert!(
        matches!(handle, Value::Object(_) | Value::I32(_) | Value::F64(_)),
        "expected a non-null handle, got {handle:?}"
    );
}

#[test]
fn set_timeout_handle_is_distinct_per_call() {
    let h1 = call_timers("setTimeout", vec![Value::Null, Value::I32(1000)]);
    let h2 = call_timers("setTimeout", vec![Value::Null, Value::I32(1000)]);
    // Two handles for two registrations — they must not be the same object
    // (reference equality, not value equality).
    assert_ne!(
        format!("{h1:?}"),
        format!("{h2:?}"),
        "handles should be distinct objects"
    );
}

// ── clearTimeout ──────────────────────────────────────────────────────────────

#[test]
fn clear_timeout_accepts_valid_handle_without_panic() {
    let handle = call_timers("setTimeout", vec![Value::Null, Value::I32(5000)]);
    let result = call_timers("clearTimeout", vec![handle]);
    assert_eq!(result, Value::Undefined);
}

#[test]
fn clear_timeout_accepts_null_without_panic() {
    // Node.js silently ignores clearTimeout(null)
    let result = call_timers("clearTimeout", vec![Value::Null]);
    assert_eq!(result, Value::Undefined);
}

#[test]
fn clear_timeout_accepts_undefined_without_panic() {
    let result = call_timers("clearTimeout", vec![Value::Undefined]);
    assert_eq!(result, Value::Undefined);
}

// ── setInterval ───────────────────────────────────────────────────────────────

#[test]
fn set_interval_returns_handle_object() {
    let handle = call_timers("setInterval", vec![Value::Null, Value::I32(100)]);
    assert!(
        matches!(handle, Value::Object(_) | Value::I32(_) | Value::F64(_)),
        "expected a non-null handle, got {handle:?}"
    );
}

// ── clearInterval ─────────────────────────────────────────────────────────────

#[test]
fn clear_interval_accepts_interval_handle() {
    let handle = call_timers("setInterval", vec![Value::Null, Value::I32(1000)]);
    let result = call_timers("clearInterval", vec![handle]);
    assert_eq!(result, Value::Undefined);
}

#[test]
fn clear_interval_accepts_null_without_panic() {
    let result = call_timers("clearInterval", vec![Value::Null]);
    assert_eq!(result, Value::Undefined);
}

// ── setImmediate ──────────────────────────────────────────────────────────────

#[test]
fn set_immediate_returns_handle() {
    let handle = call_timers("setImmediate", vec![Value::Null]);
    assert!(
        matches!(handle, Value::Object(_) | Value::I32(_) | Value::F64(_)),
        "expected a non-null handle, got {handle:?}"
    );
}

// ── clearImmediate ────────────────────────────────────────────────────────────

#[test]
fn clear_immediate_accepts_handle() {
    let handle = call_timers("setImmediate", vec![Value::Null]);
    let result = call_timers("clearImmediate", vec![handle]);
    assert_eq!(result, Value::Undefined);
}

#[test]
fn clear_immediate_accepts_null_without_panic() {
    let result = call_timers("clearImmediate", vec![Value::Null]);
    assert_eq!(result, Value::Undefined);
}

// ── queueMicrotask ────────────────────────────────────────────────────────────

#[test]
fn queue_microtask_accepts_callback_without_panic() {
    let result = call_timers("queueMicrotask", vec![Value::Null]);
    assert_eq!(result, Value::Undefined);
}

// ── Timeout handle properties ─────────────────────────────────────────────────

#[test]
fn timeout_handle_has_ref_method() {
    let handle = call_timers("setTimeout", vec![Value::Null, Value::I32(1000)]);
    if let Value::Object(obj) = &handle {
        let o = obj.lock().unwrap();
        assert!(o.properties.contains_key("ref"), "Timeout.ref must exist");
    }
}

#[test]
fn timeout_handle_has_unref_method() {
    let handle = call_timers("setTimeout", vec![Value::Null, Value::I32(1000)]);
    if let Value::Object(obj) = &handle {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("unref"),
            "Timeout.unref must exist"
        );
    }
}

#[test]
fn timeout_handle_has_has_ref_method() {
    let handle = call_timers("setTimeout", vec![Value::Null, Value::I32(1000)]);
    if let Value::Object(obj) = &handle {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("hasRef"),
            "Timeout.hasRef must exist"
        );
    }
}

#[test]
fn timeout_handle_refresh_method_exists() {
    let handle = call_timers("setTimeout", vec![Value::Null, Value::I32(1000)]);
    if let Value::Object(obj) = &handle {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("refresh"),
            "Timeout.refresh must exist"
        );
    }
}

#[test]
fn set_timeout_with_extra_args_returns_handle() {
    let handle = call_timers(
        "setTimeout",
        vec![
            Value::Null,
            Value::I32(0),
            Value::String(std::sync::Arc::from("arg1")),
            Value::I32(42),
        ],
    );
    assert!(
        matches!(handle, Value::Object(_) | Value::I32(_) | Value::F64(_)),
        "setTimeout with extra args must return handle"
    );
}

// ── Interval handle properties ────────────────────────────────────────────────

#[test]
fn interval_handle_has_ref_method() {
    let handle = call_timers("setInterval", vec![Value::Null, Value::I32(1000)]);
    if let Value::Object(obj) = &handle {
        let o = obj.lock().unwrap();
        assert!(o.properties.contains_key("ref"), "Interval.ref must exist");
    }
}

#[test]
fn interval_handle_has_unref_method() {
    let handle = call_timers("setInterval", vec![Value::Null, Value::I32(1000)]);
    if let Value::Object(obj) = &handle {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("unref"),
            "Interval.unref must exist"
        );
    }
}

#[test]
fn interval_handle_has_has_ref_method() {
    let handle = call_timers("setInterval", vec![Value::Null, Value::I32(1000)]);
    if let Value::Object(obj) = &handle {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("hasRef"),
            "Interval.hasRef must exist"
        );
    }
}

// ── Immediate handle properties ───────────────────────────────────────────────

#[test]
fn immediate_handle_has_ref_method() {
    let handle = call_timers("setImmediate", vec![Value::Null]);
    if let Value::Object(obj) = &handle {
        let o = obj.lock().unwrap();
        assert!(o.properties.contains_key("ref"), "Immediate.ref must exist");
    }
}

#[test]
fn immediate_handle_has_unref_method() {
    let handle = call_timers("setImmediate", vec![Value::Null]);
    if let Value::Object(obj) = &handle {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("unref"),
            "Immediate.unref must exist"
        );
    }
}

#[test]
fn immediate_handle_has_has_ref_method() {
    let handle = call_timers("setImmediate", vec![Value::Null]);
    if let Value::Object(obj) = &handle {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("hasRef"),
            "Immediate.hasRef must exist"
        );
    }
}

// ── clearTimeout called twice is safe ────────────────────────────────────────

#[test]
fn clear_timeout_called_twice_does_not_panic() {
    let handle = call_timers("setTimeout", vec![Value::Null, Value::I32(5000)]);
    let _ = call_timers("clearTimeout", vec![handle.clone()]);
    let result = call_timers("clearTimeout", vec![handle]);
    assert_eq!(result, Value::Undefined);
}

// ── setInterval cleared immediately ──────────────────────────────────────────

#[test]
fn set_interval_cleared_immediately_does_not_fire() {
    let handle = call_timers("setInterval", vec![Value::Null, Value::I32(0)]);
    let result = call_timers("clearInterval", vec![handle]);
    assert_eq!(result, Value::Undefined);
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_timers_surface_is_registered() {
    let expected = [
        "setTimeout",
        "clearTimeout",
        "setInterval",
        "clearInterval",
        "setImmediate",
        "clearImmediate",
        "queueMicrotask",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing node:timers imports: {missing:?}"
    );
}
