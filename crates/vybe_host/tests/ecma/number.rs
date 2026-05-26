//! Behaviour tests for `ecma:number` host imports — the ECMA-262
//! `Number` static surface, `Number.prototype`, plus the global
//! coercion/predicate functions (§19.2.4 isNaN, isFinite, §19.2.5
//! parseInt, parseFloat).
//!
//! Reference: ECMA-262 §21.1 Number.
//!
//! Where the merged `wasm:js-number` proposal covers a primitive
//! op (`test`, `testI32`, `fromF64`, `fromI32`, `toF64`, `toI32`),
//! `ecma:number` defers to it; everything spec-defined that lives
//! beyond that proposal (parseInt/parseFloat/isNaN/isFinite/etc.)
//! is registered here.

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-number-test>");
    let import_idx = chunk.add_import("ecma:number", name);
    let argc = args.len() as u8;
    for value in args {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

fn as_string(value: &Value) -> String {
    match value {
        Value::String(text) => text.to_string(),
        other => format!("{}", other),
    }
}

// ── Constants (Number.MAX_SAFE_INTEGER etc. — exposed as 0-arg fns) ─

#[test]
fn max_safe_integer_is_2_pow_53_minus_1() {
    assert_eq!(invoke("MAX_SAFE_INTEGER", vec![]), Value::F64(9007199254740991.0));
}

#[test]
fn min_safe_integer_is_negative_2_pow_53_minus_1() {
    assert_eq!(invoke("MIN_SAFE_INTEGER", vec![]), Value::F64(-9007199254740991.0));
}

#[test]
fn epsilon_constant() {
    let v = invoke("EPSILON", vec![]);
    if let Value::F64(eps) = v {
        // EPSILON ≈ 2.220446049250313e-16; check the order of magnitude.
        assert!(eps > 0.0 && eps < 1e-15, "EPSILON should be tiny positive, got {}", eps);
    } else {
        panic!("EPSILON expected f64, got {:?}", v);
    }
}

#[test]
fn positive_infinity_constant() {
    if let Value::F64(n) = invoke("POSITIVE_INFINITY", vec![]) {
        assert!(n.is_infinite() && n.is_sign_positive());
    } else {
        panic!("POSITIVE_INFINITY expected f64");
    }
}

#[test]
fn negative_infinity_constant() {
    if let Value::F64(n) = invoke("NEGATIVE_INFINITY", vec![]) {
        assert!(n.is_infinite() && n.is_sign_negative());
    } else {
        panic!("NEGATIVE_INFINITY expected f64");
    }
}

#[test]
fn nan_constant_is_nan() {
    if let Value::F64(n) = invoke("NaN", vec![]) {
        assert!(n.is_nan());
    } else {
        panic!("NaN expected f64");
    }
}

// ── Number.isFinite / Number.isNaN — STRICT (no coercion) ─────────

#[test]
fn is_finite_true_for_finite_numbers() {
    assert_eq!(invoke("isFinite", vec![Value::F64(42.0)]), Value::Bool(true));
    assert_eq!(invoke("isFinite", vec![Value::F64(0.0)]), Value::Bool(true));
}

#[test]
fn is_finite_false_for_infinities_and_nan() {
    assert_eq!(invoke("isFinite", vec![Value::F64(f64::INFINITY)]), Value::Bool(false));
    assert_eq!(invoke("isFinite", vec![Value::F64(f64::NEG_INFINITY)]), Value::Bool(false));
    assert_eq!(invoke("isFinite", vec![Value::F64(f64::NAN)]), Value::Bool(false));
}

#[test]
fn is_finite_false_for_strings() {
    // ECMA-262 §21.1.2.2 Number.isFinite: returns false for non-Number.
    assert_eq!(invoke("isFinite", vec![s("42")]), Value::Bool(false));
}

#[test]
fn is_nan_true_only_for_nan() {
    assert_eq!(invoke("isNaN", vec![Value::F64(f64::NAN)]), Value::Bool(true));
    assert_eq!(invoke("isNaN", vec![Value::F64(42.0)]), Value::Bool(false));
}

#[test]
fn is_nan_strict_false_for_strings_unlike_global() {
    // ECMA-262 §21.1.2.4 Number.isNaN: returns false for non-Number;
    // the GLOBAL isNaN coerces and would return true here.
    assert_eq!(invoke("isNaN", vec![s("hello")]), Value::Bool(false));
}

// ── Number.isInteger / isSafeInteger ──────────────────────────────

#[test]
fn is_integer_true_for_whole_numbers() {
    assert_eq!(invoke("isInteger", vec![Value::F64(42.0)]), Value::Bool(true));
    assert_eq!(invoke("isInteger", vec![Value::F64(-3.0)]), Value::Bool(true));
}

#[test]
fn is_integer_false_for_non_whole_or_non_finite() {
    assert_eq!(invoke("isInteger", vec![Value::F64(3.14)]), Value::Bool(false));
    assert_eq!(invoke("isInteger", vec![Value::F64(f64::NAN)]), Value::Bool(false));
    assert_eq!(invoke("isInteger", vec![Value::F64(f64::INFINITY)]), Value::Bool(false));
}

#[test]
fn is_safe_integer_within_range() {
    assert_eq!(invoke("isSafeInteger", vec![Value::F64(9007199254740991.0)]), Value::Bool(true));
    assert_eq!(invoke("isSafeInteger", vec![Value::F64(9007199254740992.0)]), Value::Bool(false));
}

// ── Number.parseInt / parseFloat — same as global ─────────────────

#[test]
fn parse_int_decimal_default() {
    assert_eq!(invoke("parseInt", vec![s("42")]), Value::F64(42.0));
}

#[test]
fn parse_int_with_radix() {
    assert_eq!(invoke("parseInt", vec![s("ff"), Value::F64(16.0)]), Value::F64(255.0));
    assert_eq!(invoke("parseInt", vec![s("101"), Value::F64(2.0)]), Value::F64(5.0));
}

#[test]
fn parse_int_stops_at_invalid_char() {
    // ECMA-262 §19.2.5: parseInt stops parsing at the first
    // unrecognised character, returns what was parsed so far.
    assert_eq!(invoke("parseInt", vec![s("10abc")]), Value::F64(10.0));
}

#[test]
fn parse_int_returns_nan_for_unparseable() {
    if let Value::F64(n) = invoke("parseInt", vec![s("hello")]) {
        assert!(n.is_nan(), "expected NaN, got {}", n);
    } else {
        panic!("parseInt of bad input should return NaN");
    }
}

#[test]
fn parse_float_basic() {
    assert_eq!(invoke("parseFloat", vec![s("3.14")]), Value::F64(3.14));
}

#[test]
fn parse_float_handles_exponent() {
    assert_eq!(invoke("parseFloat", vec![s("1e3")]), Value::F64(1000.0));
}

#[test]
fn parse_float_returns_nan_for_unparseable() {
    if let Value::F64(n) = invoke("parseFloat", vec![s("hello")]) {
        assert!(n.is_nan(), "expected NaN, got {}", n);
    } else {
        panic!("parseFloat of bad input should return NaN");
    }
}

// ── Number.prototype.toFixed / toString ──────────────────────────

#[test]
fn to_fixed_default_digits_zero() {
    assert_eq!(as_string(&invoke("toFixed", vec![Value::F64(3.7)])), "4");
}

#[test]
fn to_fixed_with_n_digits() {
    assert_eq!(as_string(&invoke("toFixed", vec![Value::F64(3.14159), Value::F64(2.0)])), "3.14");
}

#[test]
fn to_string_radix_decimal_default() {
    assert_eq!(as_string(&invoke("toString", vec![Value::F64(42.0)])), "42");
}

#[test]
fn to_string_radix_hex() {
    assert_eq!(as_string(&invoke("toString", vec![Value::F64(255.0), Value::F64(16.0)])), "ff");
}

#[test]
fn to_string_radix_binary() {
    assert_eq!(as_string(&invoke("toString", vec![Value::F64(5.0), Value::F64(2.0)])), "101");
}
