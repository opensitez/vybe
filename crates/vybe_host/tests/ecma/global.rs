//! Behaviour tests for `ecma:global` host imports.
//!
//! Reference: ECMA-262 §19 (The Global Object) and §18 (Global Functions).
//!
//! Covers the global coercing predicates (isNaN, isFinite) which differ from
//! their Number.* counterparts by performing ToNumber coercion first, plus
//! eval and the URI codec functions.

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-global-test>");
    let import_idx = chunk.add_import("ecma:global", name);
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

// ── global isNaN — COERCES to number first ────────────────────────────────────

#[test]
fn global_is_nan_coerces_string_to_number_and_returns_true_for_non_numeric() {
    // ECMA-262 §18.2.3: global isNaN("hello") → ToNumber("hello") = NaN → true.
    // This is the KEY difference from Number.isNaN which would return false.
    assert_eq!(invoke("isNaN", vec![s("hello")]), Value::Bool(true));
}

#[test]
fn global_is_nan_coerces_numeric_string_to_false() {
    // ToNumber("42") = 42, which is not NaN.
    assert_eq!(invoke("isNaN", vec![s("42")]), Value::Bool(false));
}

#[test]
fn global_is_nan_true_for_nan_value() {
    assert_eq!(invoke("isNaN", vec![Value::F64(f64::NAN)]), Value::Bool(true));
}

#[test]
fn global_is_nan_false_for_infinity() {
    // Infinity is not NaN.
    assert_eq!(invoke("isNaN", vec![Value::F64(f64::INFINITY)]), Value::Bool(false));
}

// ── global isFinite — COERCES to number first ─────────────────────────────────

#[test]
fn global_is_finite_coerces_numeric_string_to_true() {
    // ECMA-262 §18.2.2: global isFinite("42") → ToNumber("42") = 42 → true.
    // Number.isFinite("42") would return false (no coercion).
    assert_eq!(invoke("isFinite", vec![s("42")]), Value::Bool(true));
}

#[test]
fn global_is_finite_coerces_non_numeric_string_to_false() {
    // ToNumber("abc") = NaN; NaN is not finite.
    assert_eq!(invoke("isFinite", vec![s("abc")]), Value::Bool(false));
}

#[test]
fn global_is_finite_false_for_infinity() {
    assert_eq!(invoke("isFinite", vec![Value::F64(f64::INFINITY)]), Value::Bool(false));
}

#[test]
fn global_is_finite_true_for_finite_number() {
    assert_eq!(invoke("isFinite", vec![Value::F64(42.0)]), Value::Bool(true));
}

// ── global parseInt / parseFloat (same as Number.parseInt/parseFloat) ─────────

#[test]
fn global_parse_int_decimal() {
    assert_eq!(invoke("parseInt", vec![s("42")]), Value::F64(42.0));
}

#[test]
fn global_parse_float_decimal() {
    assert_eq!(invoke("parseFloat", vec![s("3.14")]), Value::F64(3.14));
}

// ── eval — executes a string as ECMAScript code ───────────────────────────────

#[test]
fn eval_of_integer_literal_returns_integer() {
    // ECMA-262 §18.2.1: eval("42") → 42.
    assert_eq!(invoke("eval", vec![s("42")]), Value::F64(42.0));
}

#[test]
fn eval_of_non_string_returns_the_value_unchanged() {
    // ECMA-262 §18.2.1: If argument is not a String, return argument directly.
    assert_eq!(invoke("eval", vec![Value::I32(7)]), Value::I32(7));
}

// ── globalThis ────────────────────────────────────────────────────────────────

#[test]
fn global_this_is_an_object() {
    // ECMA-262 §19.1 globalThis: must be an object.
    let gt = invoke("globalThis", vec![]);
    assert!(matches!(gt, Value::Object(_)));
}

// ── Infinity and NaN as global properties ────────────────────────────────────

#[test]
fn infinity_global_property_is_positive_infinity() {
    // ECMA-262 §19.1.3: Infinity === Number.POSITIVE_INFINITY.
    if let Value::F64(v) = invoke("Infinity", vec![]) {
        assert!(v.is_infinite() && v.is_sign_positive());
    } else {
        panic!("expected F64");
    }
}

#[test]
fn nan_global_property_is_nan() {
    // ECMA-262 §19.1.4: NaN === Number.NaN.
    if let Value::F64(v) = invoke("NaN", vec![]) {
        assert!(v.is_nan());
    } else {
        panic!("expected F64");
    }
}

#[test]
fn undefined_global_property_is_undefined() {
    // ECMA-262 §19.1.5: undefined is the undefined value.
    assert_eq!(invoke("undefined", vec![]), Value::Undefined);
}
