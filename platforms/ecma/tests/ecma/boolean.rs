//! Behaviour tests for `ecma:boolean` host imports.
//!
//! Reference: ECMA-262 §20.3 Boolean.
//!
//! Covers the Boolean() coercion rules (ToBoolean), Boolean object wrapper,
//! and valueOf. Each test covers a distinct ECMA-262 behaviour.

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

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

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-boolean-test>");
    let import_idx = chunk.add_import("ecma:boolean", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

// ── ToBoolean — falsy values ─────────────────────────────────────────────────

#[test]
fn false_is_falsy() {
    assert_eq!(
        invoke("toBoolean", vec![Value::Bool(false)]),
        Value::Bool(false)
    );
}

#[test]
fn zero_is_falsy() {
    // ECMA-262 §7.1.2: +0, -0, and 0n are all falsy.
    assert_eq!(invoke("toBoolean", vec![Value::I32(0)]), Value::Bool(false));
    assert_eq!(
        invoke("toBoolean", vec![Value::F64(0.0)]),
        Value::Bool(false)
    );
}

#[test]
fn nan_is_falsy() {
    assert_eq!(
        invoke("toBoolean", vec![Value::F64(f64::NAN)]),
        Value::Bool(false)
    );
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
    assert_eq!(
        invoke("toBoolean", vec![Value::Undefined]),
        Value::Bool(false)
    );
}

// ── ToBoolean — truthy values ────────────────────────────────────────────────

#[test]
fn non_zero_number_is_truthy() {
    assert_eq!(invoke("toBoolean", vec![Value::I32(1)]), Value::Bool(true));
    assert_eq!(
        invoke("toBoolean", vec![Value::F64(-0.5)]),
        Value::Bool(true)
    );
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
    use vybe_runtime::value::Object;
    let empty_obj = Value::Object(Arc::new(Mutex::new(Object::new())));
    assert_eq!(invoke("toBoolean", vec![empty_obj]), Value::Bool(true));
}

#[test]
fn positive_infinity_is_truthy() {
    assert_eq!(
        invoke("toBoolean", vec![Value::F64(f64::INFINITY)]),
        Value::Bool(true)
    );
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
    assert_eq!(
        invoke("valueOf", vec![Value::Bool(true)]),
        Value::Bool(true)
    );
}

#[test]
fn value_of_false_is_false() {
    assert_eq!(
        invoke("valueOf", vec![Value::Bool(false)]),
        Value::Bool(false)
    );
}
