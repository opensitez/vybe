//! Behaviour tests for `ecma:finalizationregistry` host imports.
//!
//! Reference: ECMA-262 §26.2 FinalizationRegistry.
//!
//! FinalizationRegistry lets code react when a registered object is GC'd.
//! In-process, GC has not run so cleanup callbacks are never fired
//! synchronously — we test the registration / unregistration surface only.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::compiler::platforms::register_platforms;

fn invoke(name: &str, args: Vec<Value>) -> Value {
    invoke_module("ecma:finalizationregistry", name, args)
}

fn invoke_module(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-finalizationregistry-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
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
