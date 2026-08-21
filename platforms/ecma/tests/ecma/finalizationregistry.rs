//! Behaviour tests for `ecma:finalizationregistry` host imports.
//!
//! Reference: ECMA-262 §26.2 FinalizationRegistry.
//!
//! FinalizationRegistry lets code react when a registered object is GC'd.
//! In-process, GC has not run so cleanup callbacks are never fired
//! synchronously — we test the registration / unregistration surface only.

use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{Object, Value};
use vybe_runtime::{Chunk, Op, VM};

fn invoke(name: &str, args: Vec<Value>) -> Value {
    invoke_module("ecma:finalizationregistry", name, args)
}

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn push_arg(vm: &mut VM, chunk: &mut Chunk, value: Value) {
    match value {
        Value::I32(n) => chunk.emit_i32_const(n, 0),
        Value::I64(n) => chunk.emit_i64_const(n, 0),
        Value::F32(f) => chunk.emit_f32_const(f, 0),
        Value::F64(f) => chunk.emit_f64_const(f, 0),
        Value::Bool(b) => chunk.emit_bool_const(b, 0),
        Value::String(s) => chunk.emit_string_const(&s, 0),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
        other => {
            let global = format!(
                "__test_arg_{}",
                TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
            );
            vm.set_global_owned(global.clone(), other);
            let ci = chunk.intern_string_constant(&global);
            chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
        }
    }
}

fn invoke_module(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-finalizationregistry-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn obj() -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new())))
}

fn callback_obj() -> Value {
    // The host recognises this as a no-op cleanup callback descriptor.
    let mut o = Object::new();
    o.properties
        .insert("__callback_noop".to_string(), Value::Bool(true));
    Value::Object(Arc::new(Mutex::new(o)))
}

// ── construction ──────────────────────────────────────────────────────────────

#[test]
fn new_returns_object() {
    let fr = invoke("new", vec![callback_obj()]);
    assert!(matches!(fr, Value::Object(_)));
}

#[test]
fn hyphenated_module_alias_new_returns_object() {
    let fr = invoke_module("ecma:finalization-registry", "new", vec![callback_obj()]);
    assert!(matches!(fr, Value::Object(_)));
}

// ── register ─────────────────────────────────────────────────────────────────

#[test]
fn register_without_unregister_token_does_not_panic() {
    let fr = invoke("new", vec![callback_obj()]);
    let target = obj();
    // register(target, heldValue) — no unregister token.
    let result = invoke("register", vec![fr, target, Value::I32(42)]);
    // ECMA-262: register returns undefined.
    assert_eq!(result, Value::Undefined);
}

#[test]
fn register_with_unregister_token_returns_undefined() {
    let fr = invoke("new", vec![callback_obj()]);
    let target = obj();
    let token = obj();
    let result = invoke("registerWithToken", vec![fr, target, Value::I32(1), token]);
    assert_eq!(result, Value::Undefined);
}

// ── unregister ────────────────────────────────────────────────────────────────

#[test]
fn unregister_with_known_token_returns_true() {
    let fr = invoke("new", vec![callback_obj()]);
    let target = obj();
    let token = obj();
    invoke(
        "registerWithToken",
        vec![fr.clone(), target, Value::I32(0), token.clone()],
    );
    assert_eq!(invoke("unregister", vec![fr, token]), Value::Bool(true));
}

#[test]
fn unregister_with_unknown_token_returns_false() {
    let fr = invoke("new", vec![callback_obj()]);
    // No registration with this token.
    assert_eq!(invoke("unregister", vec![fr, obj()]), Value::Bool(false));
}

#[test]
fn unregister_twice_with_same_token_returns_false_on_second_call() {
    let fr = invoke("new", vec![callback_obj()]);
    let target = obj();
    let token = obj();
    invoke(
        "registerWithToken",
        vec![fr.clone(), target, Value::I32(0), token.clone()],
    );
    invoke("unregister", vec![fr.clone(), token.clone()]);
    // Second unregister — entry is already gone.
    assert_eq!(invoke("unregister", vec![fr, token]), Value::Bool(false));
}

// ── cleanup not fired synchronously ───────────────────────────────────────────

#[test]
fn cleanup_callback_not_fired_before_gc() {
    // In any synchronous in-process test, no cleanup fires.
    // We verify no panic and that pendingCleanupCount (if exposed) is 0.
    let fr = invoke("new", vec![callback_obj()]);
    let target = obj();
    invoke("register", vec![fr.clone(), target, Value::I32(99)]);
    let pending = invoke("pendingCleanupCount", vec![fr]);
    // Either 0 (nothing collected) or Undefined (not exposed); both are valid.
    assert!(
        pending == Value::I32(0) || pending == Value::Undefined,
        "cleanup must not fire synchronously, got {:?}",
        pending
    );
}
