//! Behaviour tests for `ecma:boolean` host imports.
//!
//! Reference: ECMA-262 §20.3 Boolean.
//!
//! Covers the Boolean() coercion rules (ToBoolean), Boolean object wrapper,
//! and valueOf. Each test covers a distinct ECMA-262 behaviour.

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-boolean-test>");
    let import_idx = chunk.add_import("ecma:boolean", name);
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

fn s(text: &str) -> Value { Value::String(std::sync::Arc::from(text)) }

// ── ToBoolean — falsy values ─────────────────────────────────────────────────

#[test]
fn false_is_falsy() {
    assert_eq!(invoke("toBoolean", vec![Value::Bool(false)]), Value::Bool(false));
}

#[test]
fn zero_is_falsy() {
    // ECMA-262 §7.1.2: +0, -0, and 0n are all falsy.
    assert_eq!(invoke("toBoolean", vec![Value::I32(0)]), Value::Bool(false));
    assert_eq!(invoke("toBoolean", vec![Value::F64(0.0)]), Value::Bool(false));
}

#[test]
fn nan_is_falsy() {
    assert_eq!(invoke("toBoolean", vec![Value::F64(f64::NAN)]), Value::Bool(false));
}

#[test]
fn empty_string_is_falsy() {
    assert_eq!(invoke("toBoolean", vec![s("")]), Value::Bool(false));
}

#[test]
fn null_is_falsy() {
    assert_eq!(invoke("toBoolean", vec![Value::Null]), Value::Bool(false));
}

#[test]
fn undefined_is_falsy() {
    assert_eq!(invoke("toBoolean", vec![Value::Undefined]), Value::Bool(false));
}

// ── ToBoolean — truthy values ────────────────────────────────────────────────

#[test]
fn non_zero_number_is_truthy() {
    assert_eq!(invoke("toBoolean", vec![Value::I32(1)]), Value::Bool(true));
    assert_eq!(invoke("toBoolean", vec![Value::F64(-0.5)]), Value::Bool(true));
}

#[test]
fn non_empty_string_is_truthy() {
    assert_eq!(invoke("toBoolean", vec![s("false")]), Value::Bool(true));
    assert_eq!(invoke("toBoolean", vec![s(" ")]), Value::Bool(true));
}

#[test]
fn string_zero_is_truthy() {
    // ECMA-262: the string "0" is truthy (only the NUMBER 0 is falsy).
    assert_eq!(invoke("toBoolean", vec![s("0")]), Value::Bool(true));
}

#[test]
fn object_is_always_truthy() {
    // ECMA-262 §7.1.2: all objects (including empty ones) are truthy.
    use std::sync::{Arc, Mutex};
    use vybe_bytecode::value::Object;
    let empty_obj = Value::Object(Arc::new(Mutex::new(Object::new())));
    assert_eq!(invoke("toBoolean", vec![empty_obj]), Value::Bool(true));
}

#[test]
fn positive_infinity_is_truthy() {
    assert_eq!(invoke("toBoolean", vec![Value::F64(f64::INFINITY)]), Value::Bool(true));
}

// ── Boolean() constructor (primitive wrapper) ─────────────────────────────────

#[test]
fn boolean_constructor_with_true_returns_true_wrapper() {
    let b = invoke("new", vec![Value::Bool(true)]);
    // valueOf on the wrapper must return the primitive.
    assert_eq!(invoke("valueOf", vec![b]), Value::Bool(true));
}

#[test]
fn boolean_constructor_with_false_returns_false_wrapper() {
    let b = invoke("new", vec![Value::Bool(false)]);
    assert_eq!(invoke("valueOf", vec![b]), Value::Bool(false));
}

#[test]
fn boolean_constructor_coerces_zero_to_false() {
    let b = invoke("new", vec![Value::I32(0)]);
    assert_eq!(invoke("valueOf", vec![b]), Value::Bool(false));
}

#[test]
fn boolean_constructor_coerces_one_to_true() {
    let b = invoke("new", vec![Value::I32(1)]);
    assert_eq!(invoke("valueOf", vec![b]), Value::Bool(true));
}

// ── Boolean.prototype.toString ────────────────────────────────────────────────

#[test]
fn to_string_of_true_is_string_true() {
    assert_eq!(invoke("toString", vec![Value::Bool(true)]), s("true"));
}

#[test]
fn to_string_of_false_is_string_false() {
    assert_eq!(invoke("toString", vec![Value::Bool(false)]), s("false"));
}

// ── valueOf round-trips ────────────────────────────────────────────────────────

#[test]
fn value_of_true_is_true() {
    assert_eq!(invoke("valueOf", vec![Value::Bool(true)]), Value::Bool(true));
}

#[test]
fn value_of_false_is_false() {
    assert_eq!(invoke("valueOf", vec![Value::Bool(false)]), Value::Bool(false));
}
