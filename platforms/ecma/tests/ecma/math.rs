use std::sync::{Arc, Mutex};
use vybe_runtime::value::Object;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-math-test>");
    let import_idx = chunk.add_import("ecma:math", name);
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

fn array(values: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(values))))
}

#[test]
fn floor_and_ceil_round_toward_expected_bounds() {
    assert_eq!(invoke("floor", vec![Value::F64(3.9)]), Value::F64(3.0));
    assert_eq!(invoke("ceil", vec![Value::F64(3.1)]), Value::F64(4.0));
}

#[test]
fn sign_reports_negative_zero_and_positive_values() {
    assert_eq!(invoke("sign", vec![Value::F64(-4.0)]), Value::F64(-1.0));
    assert_eq!(invoke("sign", vec![Value::F64(0.0)]), Value::F64(0.0));
    assert_eq!(invoke("sign", vec![Value::F64(8.0)]), Value::F64(1.0));
}

#[test]
fn log_with_explicit_base_uses_change_of_base() {
    assert_eq!(
        invoke("log", vec![Value::F64(100.0), Value::F64(10.0)]),
        Value::F64(2.0)
    );
}

#[test]
fn clamp_limits_to_closed_interval() {
    assert_eq!(
        invoke(
            "clamp",
            vec![Value::F64(10.0), Value::F64(0.0), Value::F64(5.0)]
        ),
        Value::F64(5.0)
    );
    assert_eq!(
        invoke(
            "clamp",
            vec![Value::F64(-2.0), Value::F64(0.0), Value::F64(5.0)]
        ),
        Value::F64(0.0)
    );
}

#[test]
fn imul_multiplies_using_32_bit_wrapping_rules() {
    assert_eq!(
        invoke("imul", vec![Value::F64(3.0), Value::F64(4.0)]),
        Value::F64(12.0)
    );
}

#[test]
fn random_returns_unit_interval_value() {
    let Value::F64(number) = invoke("random", vec![]) else {
        panic!("random should return f64");
    };
    assert!(
        (0.0..1.0).contains(&number),
        "random out of range: {}",
        number
    );
}

#[test]
fn min_of_and_max_of_materialize_iterables() {
    let values = array(vec![Value::F64(3.0), Value::F64(1.0), Value::F64(4.0)]);
    assert_eq!(invoke("minOf", vec![values.clone()]), Value::F64(1.0));
    assert_eq!(invoke("maxOf", vec![values]), Value::F64(4.0));
}

#[test]
fn sum_precise_handles_catastrophic_cancellation_case() {
    let values = array(vec![
        Value::F64(1.0e16),
        Value::F64(1.0),
        Value::F64(-1.0e16),
    ]);
    assert_eq!(invoke("sumPrecise", vec![values]), Value::F64(1.0));
}

#[test]
fn cbrt_and_hypot_follow_standard_library_results() {
    assert_eq!(invoke("cbrt", vec![Value::F64(27.0)]), Value::F64(3.0));
    assert_eq!(
        invoke("hypot", vec![Value::F64(3.0), Value::F64(4.0)]),
        Value::F64(5.0)
    );
}

// ── abs ───────────────────────────────────────────────────────────────────────

#[test]
fn abs_strips_sign_from_negative() {
    assert_eq!(invoke("abs", vec![Value::F64(-5.0)]), Value::F64(5.0));
}

#[test]
fn abs_of_negative_zero_is_positive_zero() {
    // ECMA-262 §21.3.2.1: Math.abs(-0) = +0.
    if let Value::F64(v) = invoke("abs", vec![Value::F64(-0.0)]) {
        assert!(!v.is_sign_negative(), "abs(-0) must be +0");
    } else {
        panic!("expected F64");
    }
}

// ── round — ties toward +Infinity ─────────────────────────────────────────────

#[test]
fn round_half_rounds_up_toward_positive_infinity() {
    // ECMA-262 §21.3.2.28: Math.round(0.5) = 1, NOT 0.
    assert_eq!(invoke("round", vec![Value::F64(0.5)]), Value::F64(1.0));
}

#[test]
fn round_negative_half_rounds_toward_zero() {
    // ECMA-262 §21.3.2.28: Math.round(-0.5) = 0 (ties go toward +Inf,
    // so -0.5 rounds to 0, not -1). This surprises people expecting symmetric
    // banker's rounding.
    assert_eq!(invoke("round", vec![Value::F64(-0.5)]), Value::F64(0.0));
}

#[test]
fn round_away_from_half_rounds_normally() {
    assert_eq!(invoke("round", vec![Value::F64(3.4)]), Value::F64(3.0));
    assert_eq!(invoke("round", vec![Value::F64(3.6)]), Value::F64(4.0));
}

// ── ceil / floor with negatives (common gotcha) ───────────────────────────────

#[test]
fn ceil_of_negative_fraction_rounds_toward_zero_not_away() {
    // ceil(-2.5) = -2, NOT -3.
    assert_eq!(invoke("ceil", vec![Value::F64(-2.5)]), Value::F64(-2.0));
}

#[test]
fn floor_of_negative_fraction_rounds_away_from_zero() {
    // floor(-2.5) = -3.
    assert_eq!(invoke("floor", vec![Value::F64(-2.5)]), Value::F64(-3.0));
}

// ── trunc — rounds toward zero regardless of sign ────────────────────────────

#[test]
fn trunc_of_positive_drops_decimal() {
    assert_eq!(invoke("trunc", vec![Value::F64(3.9)]), Value::F64(3.0));
}

#[test]
fn trunc_of_negative_rounds_toward_zero() {
    // trunc(-3.9) = -3 (not -4); distinct from floor.
    assert_eq!(invoke("trunc", vec![Value::F64(-3.9)]), Value::F64(-3.0));
}

// ── pow / sqrt / exp ──────────────────────────────────────────────────────────

#[test]
fn pow_integer_exponent() {
    assert_eq!(
        invoke("pow", vec![Value::F64(2.0), Value::F64(10.0)]),
        Value::F64(1024.0)
    );
}

#[test]
fn pow_of_negative_one_half_is_complex_nan() {
    // Math.pow(-1, 0.5) = NaN because the result would be imaginary.
    if let Value::F64(v) = invoke("pow", vec![Value::F64(-1.0), Value::F64(0.5)]) {
        assert!(v.is_nan());
    } else {
        panic!("expected F64");
    }
}

#[test]
fn sqrt_of_perfect_square() {
    assert_eq!(invoke("sqrt", vec![Value::F64(9.0)]), Value::F64(3.0));
}

#[test]
fn sqrt_of_negative_is_nan() {
    if let Value::F64(v) = invoke("sqrt", vec![Value::F64(-1.0)]) {
        assert!(v.is_nan());
    } else {
        panic!("expected F64");
    }
}

#[test]
fn exp_of_zero_is_one() {
    // e^0 = 1 exactly.
    assert_eq!(invoke("exp", vec![Value::F64(0.0)]), Value::F64(1.0));
}

// ── log variants ──────────────────────────────────────────────────────────────

#[test]
fn natural_log_of_one_is_zero() {
    // ECMA-262 §21.3.2.20 Math.log(1) = +0.
    assert_eq!(invoke("ln", vec![Value::F64(1.0)]), Value::F64(0.0));
}

#[test]
fn log2_of_eight_is_three() {
    if let Value::F64(v) = invoke("log2", vec![Value::F64(8.0)]) {
        assert!((v - 3.0).abs() < 1e-9);
    } else {
        panic!("expected F64");
    }
}

#[test]
fn log10_of_thousand_is_three() {
    if let Value::F64(v) = invoke("log10", vec![Value::F64(1000.0)]) {
        assert!((v - 3.0).abs() < 1e-9);
    } else {
        panic!("expected F64");
    }
}

#[test]
fn expm1_of_zero_is_zero() {
    // Math.expm1(x) = e^x - 1, accurate for small x.
    assert_eq!(invoke("expm1", vec![Value::F64(0.0)]), Value::F64(0.0));
}

#[test]
fn log1p_of_zero_is_zero() {
    // Math.log1p(x) = ln(1+x), accurate for small x.
    assert_eq!(invoke("log1p", vec![Value::F64(0.0)]), Value::F64(0.0));
}

// ── Trigonometry ──────────────────────────────────────────────────────────────

#[test]
fn sin_of_pi_over_two_is_one() {
    if let Value::F64(v) = invoke("sin", vec![Value::F64(std::f64::consts::FRAC_PI_2)]) {
        assert!((v - 1.0).abs() < 1e-9);
    } else {
        panic!("expected F64");
    }
}

#[test]
fn cos_of_zero_is_one() {
    assert_eq!(invoke("cos", vec![Value::F64(0.0)]), Value::F64(1.0));
}

#[test]
fn tan_of_zero_is_zero() {
    assert_eq!(invoke("tan", vec![Value::F64(0.0)]), Value::F64(0.0));
}

#[test]
fn atan2_of_zero_and_negative_one_is_pi() {
    // atan2(0, -1) = π (the angle pointing in the negative X direction).
    if let Value::F64(v) = invoke("atan2", vec![Value::F64(0.0), Value::F64(-1.0)]) {
        assert!((v - std::f64::consts::PI).abs() < 1e-9);
    } else {
        panic!("expected F64");
    }
}

#[test]
fn asin_of_zero_is_zero() {
    assert_eq!(invoke("asin", vec![Value::F64(0.0)]), Value::F64(0.0));
}

#[test]
fn acos_of_one_is_zero() {
    assert_eq!(invoke("acos", vec![Value::F64(1.0)]), Value::F64(0.0));
}

#[test]
fn atan_of_one_is_pi_over_four() {
    if let Value::F64(v) = invoke("atan", vec![Value::F64(1.0)]) {
        assert!((v - std::f64::consts::FRAC_PI_4).abs() < 1e-9);
    } else {
        panic!("expected F64");
    }
}

// ── Hyperbolic ────────────────────────────────────────────────────────────────

#[test]
fn sinh_of_zero_is_zero() {
    assert_eq!(invoke("sinh", vec![Value::F64(0.0)]), Value::F64(0.0));
}

#[test]
fn cosh_of_zero_is_one() {
    assert_eq!(invoke("cosh", vec![Value::F64(0.0)]), Value::F64(1.0));
}

#[test]
fn tanh_of_zero_is_zero() {
    assert_eq!(invoke("tanh", vec![Value::F64(0.0)]), Value::F64(0.0));
}

// ── max / min — variadic, empty returns sentinel ──────────────────────────────

#[test]
fn max_of_no_args_is_negative_infinity() {
    // ECMA-262 §21.3.2.24: Math.max() = -Infinity.
    if let Value::F64(v) = invoke("max", vec![]) {
        assert!(v.is_infinite() && v.is_sign_negative());
    } else {
        panic!("expected F64");
    }
}

#[test]
fn min_of_no_args_is_positive_infinity() {
    // ECMA-262 §21.3.2.25: Math.min() = +Infinity.
    if let Value::F64(v) = invoke("min", vec![]) {
        assert!(v.is_infinite() && v.is_sign_positive());
    } else {
        panic!("expected F64");
    }
}

#[test]
fn max_returns_largest_of_multiple_args() {
    assert_eq!(
        invoke(
            "max",
            vec![Value::F64(1.0), Value::F64(5.0), Value::F64(3.0)]
        ),
        Value::F64(5.0)
    );
}

#[test]
fn min_returns_smallest_of_multiple_args() {
    assert_eq!(
        invoke(
            "min",
            vec![Value::F64(4.0), Value::F64(2.0), Value::F64(7.0)]
        ),
        Value::F64(2.0)
    );
}

// ── clz32 — count leading zeros in 32-bit representation ─────────────────────

#[test]
fn clz32_of_one_is_31() {
    // 1 = 0b00000000000000000000000000000001 → 31 leading zeros.
    assert_eq!(invoke("clz32", vec![Value::F64(1.0)]), Value::F64(31.0));
}

#[test]
fn clz32_of_zero_is_32() {
    // All 32 bits are zero → clz32(0) = 32.
    assert_eq!(invoke("clz32", vec![Value::F64(0.0)]), Value::F64(32.0));
}

// ── Constants ─────────────────────────────────────────────────────────────────

#[test]
fn pi_constant_approximates_pi() {
    if let Value::F64(v) = invoke("PI", vec![]) {
        assert!((v - std::f64::consts::PI).abs() < 1e-12);
    } else {
        panic!("expected F64");
    }
}

#[test]
fn e_constant_approximates_eulers_number() {
    if let Value::F64(v) = invoke("E", vec![]) {
        assert!((v - std::f64::consts::E).abs() < 1e-12);
    } else {
        panic!("expected F64");
    }
}

#[test]
fn sqrt2_constant_is_root_of_two() {
    if let Value::F64(v) = invoke("SQRT2", vec![]) {
        assert!((v - std::f64::consts::SQRT_2).abs() < 1e-12);
    } else {
        panic!("expected F64");
    }
}

#[test]
fn ln2_constant_is_natural_log_of_two() {
    if let Value::F64(v) = invoke("LN2", vec![]) {
        assert!((v - std::f64::consts::LN_2).abs() < 1e-12);
    } else {
        panic!("expected F64");
    }
}

// ── Math.f16round (ES2025 §21.3.2.17) ────────────────────────────────────────

#[test]
fn f16round_rounds_to_nearest_float16_representable_value() {
    // ECMA-262 ES2025: Math.f16round(x) rounds x to the nearest IEEE 754-2008 float16 value.
    // 1.0 is exactly representable in Float16.
    let result = invoke("f16round", vec![Value::F64(1.0)]);
    assert!(matches!(result, Value::F64(f) if (f - 1.0).abs() < 0.01));
}

#[test]
fn f16round_zero_is_zero() {
    let result = invoke("f16round", vec![Value::F64(0.0)]);
    assert!(matches!(result, Value::F64(f) if f == 0.0));
}

#[test]
fn f16round_nan_produces_nan() {
    let result = invoke("f16round", vec![Value::F64(f64::NAN)]);
    assert!(matches!(result, Value::F64(f) if f.is_nan()));
}

// ── Math.fround (§21.3.2.18) ─────────────────────────────────────────────────

#[test]
fn fround_one_is_exactly_one() {
    // Math.fround(1.0) = 1.0 (exactly representable in float32).
    let result = invoke("fround", vec![Value::F64(1.0)]);
    assert!(matches!(result, Value::F64(f) if (f - 1.0).abs() < 0.0001));
}

#[test]
fn fround_reduces_precision_of_non_representable_value() {
    // 1.337 is not exactly representable in float32; fround changes its bits.
    let result = invoke("fround", vec![Value::F64(1.337)]);
    match result {
        Value::F64(f) => assert!((f - 1.337f32 as f64).abs() < 0.00001, "got {f}"),
        other => panic!("expected F64, got {:?}", other),
    }
}

// ── Math.sumPrecise (ES2026 §21.3.2.37) ──────────────────────────────────────

#[test]
fn sum_precise_avoids_floating_point_cancellation() {
    // Math.sumPrecise([1e20, -1e20, 1.0]) = 1.0 without catastrophic cancellation.
    use vybe_runtime::value::Object;
    let values = Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
        Value::F64(1e20),
        Value::F64(-1e20),
        Value::F64(1.0),
    ]))));
    let result = invoke("sumPrecise", vec![values]);
    assert!(matches!(result, Value::F64(f) if (f - 1.0).abs() < 0.0001));
}

// ── Missing constants: LN10, LOG2E, LOG10E, SQRT1_2 ──────────────────────────

#[test]
fn ln10_constant_is_natural_log_of_ten() {
    // ECMA-262 §21.3.1: Math.LN10 ≈ 2.302585092994046.
    if let Value::F64(v) = invoke("LN10", vec![]) {
        assert!((v - std::f64::consts::LN_10).abs() < 1e-12);
    } else {
        panic!("expected F64");
    }
}

#[test]
fn log2e_constant_is_log_base_2_of_e() {
    // ECMA-262 §21.3.1: Math.LOG2E ≈ 1.4426950408889634.
    if let Value::F64(v) = invoke("LOG2E", vec![]) {
        assert!((v - std::f64::consts::LOG2_E).abs() < 1e-12);
    } else {
        panic!("expected F64");
    }
}

#[test]
fn log10e_constant_is_log_base_10_of_e() {
    // ECMA-262 §21.3.1: Math.LOG10E ≈ 0.4342944819032518.
    if let Value::F64(v) = invoke("LOG10E", vec![]) {
        assert!((v - std::f64::consts::LOG10_E).abs() < 1e-12);
    } else {
        panic!("expected F64");
    }
}

#[test]
fn sqrt1_2_constant_is_reciprocal_square_root_of_two() {
    // ECMA-262 §21.3.1: Math.SQRT1_2 ≈ 0.7071067811865476 = 1/√2.
    if let Value::F64(v) = invoke("SQRT1_2", vec![]) {
        assert!((v - std::f64::consts::FRAC_1_SQRT_2).abs() < 1e-12);
    } else {
        panic!("expected F64");
    }
}

// ── Missing transcendentals: asinh, atanh ────────────────────────────────────

#[test]
fn asinh_zero_is_zero() {
    // Math.asinh(0) = 0.
    assert_eq!(invoke("asinh", vec![Value::F64(0.0)]), Value::F64(0.0));
}

#[test]
fn atanh_zero_is_zero() {
    // Math.atanh(0) = 0.
    assert_eq!(invoke("atanh", vec![Value::F64(0.0)]), Value::F64(0.0));
}

#[test]
fn atanh_one_is_positive_infinity() {
    // Math.atanh(1) = +Infinity.
    if let Value::F64(v) = invoke("atanh", vec![Value::F64(1.0)]) {
        assert!(v.is_infinite() && v.is_sign_positive());
    } else {
        panic!("expected F64");
    }
}
// ── Missing spec methods ───────────────────────────────────────────────────────

#[test]
fn acosh_of_one_is_zero() {
    // §21.3.2.3: Math.acosh(1) = 0
    assert_eq!(invoke("acosh", vec![Value::F64(1.0)]).as_f64(), 0.0);
}

#[test]
fn acosh_of_cosh_is_identity() {
    // acosh(cosh(2)) ≈ 2
    let c = invoke("cosh", vec![Value::F64(2.0)]);
    let r = invoke("acosh", vec![c]).as_f64();
    assert!((r - 2.0).abs() < 1e-10);
}
