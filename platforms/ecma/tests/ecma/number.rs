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
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

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
    register_platforms(&mut vm, &Capabilities::all());
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
    assert_eq!(
        invoke("MAX_SAFE_INTEGER", vec![]),
        Value::F64(9007199254740991.0)
    );
}

#[test]
fn min_safe_integer_is_negative_2_pow_53_minus_1() {
    assert_eq!(
        invoke("MIN_SAFE_INTEGER", vec![]),
        Value::F64(-9007199254740991.0)
    );
}

#[test]
fn epsilon_constant() {
    let v = invoke("EPSILON", vec![]);
    if let Value::F64(eps) = v {
        // EPSILON ≈ 2.220446049250313e-16; check the order of magnitude.
        assert!(
            eps > 0.0 && eps < 1e-15,
            "EPSILON should be tiny positive, got {}",
            eps
        );
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
    assert_eq!(
        invoke("isFinite", vec![Value::F64(42.0)]),
        Value::Bool(true)
    );
    assert_eq!(invoke("isFinite", vec![Value::F64(0.0)]), Value::Bool(true));
}

#[test]
fn is_finite_false_for_infinities_and_nan() {
    assert_eq!(
        invoke("isFinite", vec![Value::F64(f64::INFINITY)]),
        Value::Bool(false)
    );
    assert_eq!(
        invoke("isFinite", vec![Value::F64(f64::NEG_INFINITY)]),
        Value::Bool(false)
    );
    assert_eq!(
        invoke("isFinite", vec![Value::F64(f64::NAN)]),
        Value::Bool(false)
    );
}

#[test]
fn is_finite_false_for_strings() {
    // ECMA-262 §21.1.2.2 Number.isFinite: returns false for non-Number.
    assert_eq!(invoke("isFinite", vec![s("42")]), Value::Bool(false));
}

#[test]
fn is_nan_true_only_for_nan() {
    assert_eq!(
        invoke("isNaN", vec![Value::F64(f64::NAN)]),
        Value::Bool(true)
    );
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
    assert_eq!(
        invoke("isInteger", vec![Value::F64(42.0)]),
        Value::Bool(true)
    );
    assert_eq!(
        invoke("isInteger", vec![Value::F64(-3.0)]),
        Value::Bool(true)
    );
}

#[test]
fn is_integer_false_for_non_whole_or_non_finite() {
    assert_eq!(
        invoke("isInteger", vec![Value::F64(3.14)]),
        Value::Bool(false)
    );
    assert_eq!(
        invoke("isInteger", vec![Value::F64(f64::NAN)]),
        Value::Bool(false)
    );
    assert_eq!(
        invoke("isInteger", vec![Value::F64(f64::INFINITY)]),
        Value::Bool(false)
    );
}

#[test]
fn is_safe_integer_within_range() {
    assert_eq!(
        invoke("isSafeInteger", vec![Value::F64(9007199254740991.0)]),
        Value::Bool(true)
    );
    assert_eq!(
        invoke("isSafeInteger", vec![Value::F64(9007199254740992.0)]),
        Value::Bool(false)
    );
}

// ── Number.parseInt / parseFloat — same as global ─────────────────

#[test]
fn parse_int_decimal_default() {
    assert_eq!(invoke("parseInt", vec![s("42")]), Value::F64(42.0));
}

#[test]
fn parse_int_with_radix() {
    assert_eq!(
        invoke("parseInt", vec![s("ff"), Value::F64(16.0)]),
        Value::F64(255.0)
    );
    assert_eq!(
        invoke("parseInt", vec![s("101"), Value::F64(2.0)]),
        Value::F64(5.0)
    );
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
    assert_eq!(
        as_string(&invoke(
            "toFixed",
            vec![Value::F64(3.14159), Value::F64(2.0)]
        )),
        "3.14"
    );
}

#[test]
fn to_string_radix_decimal_default() {
    assert_eq!(as_string(&invoke("toString", vec![Value::F64(42.0)])), "42");
}

#[test]
fn to_string_radix_hex() {
    assert_eq!(
        as_string(&invoke(
            "toString",
            vec![Value::F64(255.0), Value::F64(16.0)]
        )),
        "ff"
    );
}

#[test]
fn to_string_radix_binary() {
    assert_eq!(
        as_string(&invoke("toString", vec![Value::F64(5.0), Value::F64(2.0)])),
        "101"
    );
}

// ── Number.prototype.toPrecision ─────────────────────────────────────────────

#[test]
fn to_precision_uses_significant_digits_not_decimal_places() {
    // toFixed(2) → "3.14"; toPrecision(3) → "3.14" only for small numbers.
    // toPrecision counts ALL significant digits, switching to exponential
    // notation when the number is too large or too small.
    assert_eq!(
        as_string(&invoke(
            "toPrecision",
            vec![Value::F64(123.456), Value::F64(5.0)]
        )),
        "123.46"
    );
}

#[test]
fn to_precision_switches_to_exponential_for_large_numbers() {
    // 123456.toPrecision(3) → "1.23e+5" (cannot fit 3 sig-figs without exp).
    let result = as_string(&invoke(
        "toPrecision",
        vec![Value::F64(123456.0), Value::F64(3.0)],
    ));
    assert!(
        result.contains('e') || result.contains('E'),
        "expected exponential notation, got {}",
        result
    );
}

// ── Number.prototype.toExponential ───────────────────────────────────────────

#[test]
fn to_exponential_formats_in_scientific_notation() {
    // (1234).toExponential(2) → "1.23e+3"
    let result = as_string(&invoke(
        "toExponential",
        vec![Value::F64(1234.0), Value::F64(2.0)],
    ));
    // The mantissa must be 1.23 and the exponent must encode +3.
    assert!(
        result.starts_with("1.23") && (result.contains("e+3") || result.contains("e+03")),
        "got {}",
        result
    );
}

#[test]
fn to_exponential_no_arg_uses_full_precision() {
    // (0.00123).toExponential() → "1.23e-3" (no rounding).
    let result = as_string(&invoke("toExponential", vec![Value::F64(0.00123)]));
    assert!(
        result.contains('e') || result.contains('E'),
        "expected exponential notation, got {}",
        result
    );
}

// ── Number.MAX_VALUE / MIN_VALUE ─────────────────────────────────────────────

#[test]
fn max_value_is_finite_and_very_large() {
    if let Value::F64(n) = invoke("MAX_VALUE", vec![]) {
        assert!(
            n.is_finite() && n > 1e300,
            "MAX_VALUE should be > 1e300, got {}",
            n
        );
    } else {
        panic!("expected F64");
    }
}

#[test]
fn min_value_is_the_smallest_positive_subnormal() {
    // ECMA-262: Number.MIN_VALUE is the smallest positive value, ≈ 5e-324.
    if let Value::F64(n) = invoke("MIN_VALUE", vec![]) {
        assert!(
            n > 0.0 && n < 1e-320,
            "MIN_VALUE should be tiny positive, got {}",
            n
        );
    } else {
        panic!("expected F64");
    }
}

// ── parseInt edge cases ───────────────────────────────────────────────────────

#[test]
fn parse_int_trims_leading_whitespace() {
    // ECMA-262 §18.2.5: parseInt trims leading StrWhiteSpaceChars.
    assert_eq!(invoke("parseInt", vec![s("  42")]), Value::F64(42.0));
}

#[test]
fn parse_int_auto_detects_hex_prefix() {
    // parseInt("0xff") → 255 without an explicit radix argument.
    assert_eq!(invoke("parseInt", vec![s("0xff")]), Value::F64(255.0));
}

#[test]
fn parse_int_handles_sign_prefix() {
    assert_eq!(invoke("parseInt", vec![s("-10")]), Value::F64(-10.0));
}

// ── Number.prototype.toLocaleString ──────────────────────────────────────────

#[test]
fn to_locale_string_returns_non_empty_string() {
    // ECMA-262 §21.1.3.4: toLocaleString returns a locale-sensitive string representation.
    let result = invoke("toLocaleString", vec![Value::F64(1234567.89)]);
    assert!(matches!(result, Value::String(ref s) if !s.is_empty()));
}

#[test]
fn to_locale_string_of_zero_contains_zero_digit() {
    let result = invoke("toLocaleString", vec![Value::F64(0.0)]);
    match result {
        Value::String(s) => assert!(s.contains('0')),
        other => panic!("expected string, got {:?}", other),
    }
}

// ── Global isFinite / isNaN (ECMA-262 §19.2.2 / §19.2.3) ────────────────────

#[test]
fn global_is_finite_coerces_string_to_number() {
    // ECMA-262 §19.2.2: global isFinite performs ToNumber first, unlike Number.isFinite.
    // isFinite("42") → ToNumber("42")=42 → true.
    assert_eq!(invoke("globalIsFinite", vec![s("42")]), Value::Bool(true));
}

#[test]
fn global_is_finite_false_for_infinity_string() {
    assert_eq!(
        invoke("globalIsFinite", vec![s("Infinity")]),
        Value::Bool(false)
    );
}

#[test]
fn global_is_finite_false_for_nan_value() {
    assert_eq!(
        invoke("globalIsFinite", vec![Value::F64(f64::NAN)]),
        Value::Bool(false)
    );
}

#[test]
fn global_is_nan_coerces_string_nan_to_true() {
    // ECMA-262 §19.2.3: global isNaN performs ToNumber first.
    // isNaN("NaN") → ToNumber("NaN")=NaN → true.
    assert_eq!(invoke("globalIsNaN", vec![s("NaN")]), Value::Bool(true));
}

#[test]
fn global_is_nan_false_for_numeric_string() {
    // isNaN("42") → ToNumber("42")=42 → false.
    assert_eq!(invoke("globalIsNaN", vec![s("42")]), Value::Bool(false));
}

#[test]
fn global_is_nan_true_for_non_numeric_string() {
    // isNaN("hello") → ToNumber("hello")=NaN → true.
    assert_eq!(invoke("globalIsNaN", vec![s("hello")]), Value::Bool(true));
}

// ── Number.prototype.valueOf (ECMA-262 §21.1.3.7) ────────────────────────────

#[test]
fn value_of_returns_the_number_primitive() {
    // §21.1.3.7: valueOf extracts the [[NumberData]] internal slot.
    assert_eq!(invoke("valueOf", vec![Value::F64(3.14)]), Value::F64(3.14));
}

#[test]
fn value_of_of_integer_returns_same_value() {
    assert_eq!(invoke("valueOf", vec![Value::I32(42)]), Value::I32(42));
}
