use crate::helpers::run_prints;

#[test]
fn test_abs_and_min_max_boundaries() {
    let out = run_prints(r#"
        fun main() {
            println(abs(-12))
            println(max(-5, -2))
            println(min(-5, -2))
        }
    "#);
    assert_eq!(out, &["12", "-2", "-5"]);
}

#[test]
fn test_abs_and_zero_behavior() {
    let out = run_prints(r#"
        fun main() {
            println(abs(0))
            println(abs(1))
            println(abs(-1))
        }
    "#);
    assert_eq!(out, &["0", "1", "1"]);
}

#[test]
fn test_pow_basic_chain() {
    let out = run_prints(r#"
        fun main() {
            val square = pow(3.0, 2.0)
            val cubic = pow(2.0, 3.0)
            println(square)
            println(cubic)
        }
    "#);
    assert_eq!(out, &["9", "8"]);
}

#[test]
fn test_pow_identity_edges() {
    let out = run_prints(r#"
        fun main() {
            println(pow(9.0, 0.0))
            println(pow(5.0, 1.0))
        }
    "#);
    assert_eq!(out, &["1", "5"]);
}

#[test]
fn test_pow_fractional_root_path() {
    let out = run_prints(r#"
        fun main() {
            println(pow(9.0, 0.5))
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_sqrt_chain_with_mul_and_div() {
    let out = run_prints(r#"
        fun main() {
            val value = sqrt(81.0)
            val half = value / 2.0
            println(value)
            println(half)
        }
    "#);
    assert_eq!(out, &["9", "4.5"]);
}

#[test]
fn test_floor_ceil_on_fractional_negative_inputs() {
    let out = run_prints(r#"
        fun main() {
            println(floor(-3.9))
            println(ceil(-3.9))
        }
    "#);
    assert_eq!(out, &["-4", "-3"]);
}

#[test]
fn test_rounding_negative_and_positive_edges() {
    let out = run_prints(r#"
        fun main() {
            println(round(2.4))
            println(round(-2.6))
            println(round(-2.0))
        }
    "#);
    assert_eq!(out, &["2", "-3", "-2"]);
}

#[test]
fn test_trig_zero_projection() {
    let out = run_prints(r#"
        fun main() {
            println(sin(0.0))
            println(cos(0.0))
            println(tan(0.0))
        }
    "#);
    assert_eq!(out, &["0", "1", "0"]);
}

#[test]
fn test_math_pipeline_with_nested_calls() {
    let out = run_prints(r#"
        fun main() {
            val score = abs(min(-12, -5) + max(2, 8))
            val amplified = score * score
            println(score)
            println(amplified)
        }
    "#);
    assert_eq!(out, &["5", "25"]);
}

#[test]
fn test_abs_on_int_minimum() {
    let out = run_prints(r#"
        fun main() {
            println(abs(Int.MIN_VALUE))
        }
    "#);
    assert_eq!(out, &["-2147483648"]);
}

#[test]
fn test_abs_on_long_minimum() {
    let out = run_prints(r#"
        fun main() {
            println(abs(Long.MIN_VALUE))
        }
    "#);
    assert_eq!(out, &["-9223372036854775808"]);
}

#[test]
fn test_max_min_idempotence_and_equality() {
    let out = run_prints(r#"
        fun main() {
            println(max(7, 7))
            println(min(7, 7))
            println(max(-3, -3))
            println(min(-3, -3))
        }
    "#);
    assert_eq!(out, &["7", "7", "-3", "-3"]);
}

#[test]
fn test_max_min_commute_arguments() {
    let out = run_prints(r#"
        fun main() {
            println(max(9, -4))
            println(min(9, -4))
            println(max(-4, 9))
            println(min(-4, 9))
        }
    "#);
    assert_eq!(out, &["9", "-4", "9", "-4"]);
}

#[test]
fn test_max_min_chain_behaviors() {
    let out = run_prints(r#"
        fun main() {
            val edge = max(min(-3, 8), min(2, -11))
            val span = min(max(9, 2), max(1, 4))
            println(edge)
            println(span)
        }
    "#);
    assert_eq!(out, &["2", "4"]);
}

#[test]
fn test_pow_zero_one_and_negative_base_sign() {
    let out = run_prints(r#"
        fun main() {
            println(pow(0.0, 0.0))
            println(pow(1.0, 9.0))
            println(pow(-3.0, 2.0))
            println(pow(-3.0, 3.0))
        }
    "#);
    assert_eq!(out, &["1", "1", "9", "-27"]);
}

#[test]
fn test_pow_fractional_exponent_for_integer_cube_root() {
    let out = run_prints(r#"
        fun main() {
            val rooted = round(pow(27.0, 1.0 / 3.0))
            println(rooted)
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_pow_nested_with_abs_and_sign() {
    let out = run_prints(r#"
        fun main() {
            val value = pow(abs(-12.0), 2.0)
            val signed = pow(-2.0, 4.0)
            println(value)
            println(signed)
        }
    "#);
    assert_eq!(out, &["144", "16"]);
}

#[test]
fn test_sqrt_zero_and_one() {
    let out = run_prints(r#"
        fun main() {
            println(sqrt(0.0))
            println(sqrt(1.0))
        }
    "#);
    assert_eq!(out, &["0", "1"]);
}

#[test]
fn test_sqrt_of_negative_inputs_is_nan() {
    let out = run_prints(r#"
        fun main() {
            val value = sqrt(-4.0)
            println(value.isNaN())
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_sqrt_of_squares_roundtrip() {
    let out = run_prints(r#"
        fun main() {
            val side = 13.0
            println(sqrt(side * side))
        }
    "#);
    assert_eq!(out, &["13"]);
}

#[test]
fn test_floor_and_ceil_for_integer_input() {
    let out = run_prints(r#"
        fun main() {
            println(floor(4.0))
            println(ceil(4.0))
        }
    "#);
    assert_eq!(out, &["4", "4"]);
}

#[test]
fn test_floor_and_ceil_negative_fractional_inputs() {
    let out = run_prints(r#"
        fun main() {
            println(floor(-2.3))
            println(ceil(-2.3))
        }
    "#);
    assert_eq!(out, &["-3", "-2"]);
}

#[test]
fn test_floor_ceil_rounding_roundtrip() {
    let out = run_prints(r#"
        fun main() {
            val value = 3.4
            println(floor(value) + ceil(value))
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_rounding_ties_away_from_zero() {
    let out = run_prints(r#"
        fun main() {
            println(round(2.5))
            println(round(-2.5))
            println(round(3.5))
        }
    "#);
    assert_eq!(out, &["3", "-3", "4"]);
}

#[test]
fn test_rounding_threshold_boundaries() {
    let out = run_prints(r#"
        fun main() {
            println(round(2.499_999))
            println(round(2.5))
            println(round(2.500_001))
            println(round(-2.500_001))
        }
    "#);
    assert_eq!(out, &["2", "3", "3", "-3"]);
}

#[test]
fn test_rounding_large_magnitude_preserves_sign() {
    let out = run_prints(r#"
        fun main() {
            println(round(123456.49))
            println(round(123456.50))
            println(round(-123456.51))
        }
    "#);
    assert_eq!(out, &["123456", "123457", "-123457"]);
}

#[test]
fn test_trig_zero_is_stable_baseline() {
    let out = run_prints(r#"
        fun main() {
            println(sin(0.0))
            println(cos(0.0))
            println(tan(0.0))
        }
    "#);
    assert_eq!(out, &["0", "1", "0"]);
}

#[test]
fn test_trig_odd_even_symmetry_properties() {
    let out = run_prints(r#"
        fun main() {
            val angle = 0.42
            println(abs(sin(angle) + sin(-angle)) < 1.0e-12)
            println(abs(cos(angle) - cos(-angle)) < 1.0e-12)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_finite_and_infinite_detection_for_division_quirks() {
    let out = run_prints(r#"
        fun main() {
            println((1.0 / 3.0).isFinite())
            println((1.0 / 0.0).isInfinite())
            println((-1.0 / 0.0).isInfinite())
            println((0.0 / 0.0).isNaN())
        }
    "#);
    assert_eq!(out, &["true", "true", "true", "true"]);
}

#[test]
fn test_nan_is_unordered_and_not_self_equal() {
    let out = run_prints(r#"
        fun main() {
            val value = 0.0 / 0.0
            println(value.isNaN())
            println(value == value)
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_abs_preserves_infinite_inputs_as_infinite() {
    let out = run_prints(r#"
        fun main() {
            val positive = abs(1.0 / 0.0)
            val negative = abs(-1.0 / 0.0)
            println(positive.isInfinite())
            println(negative.isInfinite())
            println(negative > 0.0)
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_abs_nan_input_stays_not_a_number() {
    let out = run_prints(r#"
        fun main() {
            val value = 0.0 / 0.0
            println(abs(value).isNaN())
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_math_pipeline_with_classification_checks() {
    let out = run_prints(r#"
        fun main() {
            val value = (pow(9.0, 2.0) - abs(-40.0))
            println(value)
            println(value.isNaN())
            println(round(sqrt(value) * 1000.0))
        }
    "#);
    assert_eq!(out, &["41", "false", "6403"]);
}

#[test]
fn test_math_value_coercion_clamps_low_high() {
    let out = run_prints(r#"
        fun main() {
            println(5.coerceIn(1, 3))
            println((-1).coerceAtLeast(0))
            println(10.coerceAtMost(7))
            println(4.coerceIn(1, 4))
        }
    "#);
    assert_eq!(out, &["3", "0", "7", "4"]);
}

#[test]
fn test_math_coerce_in_invalid_bounds_throws() {
    let out = run_prints(r#"
        fun main() {
            try {
                println(1.coerceIn(5, 2))
            } catch (e: IllegalArgumentException) {
                println("invalid")
            }
        }
    "#);
    assert_eq!(out, &["invalid"]);
}

#[test]
fn test_math_log_exp_roundtrip_small_delta() {
    let out = run_prints(r#"
        fun main() {
            println(round(exp(log(10.0))))
            println(round(exp(ln(2.0) * 3.0)))
        }
    "#);
    assert_eq!(out, &["10", "8"]);
}

#[test]
fn test_math_atan_and_tan_inverse_identity() {
    let out = run_prints(r#"
        fun main() {
            val angle = 0.75
            println(round((tan(atan(angle)) - angle) * 1e9))
            println(sign(0.0))
            println(sign(-5.0))
            println(sign(5.0))
        }
    "#);
    assert_eq!(out, &["0", "0.0", "-1.0", "1.0"]);
}

#[test]
fn test_math_hypot_and_round_trip() {
    let out = run_prints(r#"
        fun main() {
            println(hypot(3.0, 4.0))
            println(hypot(0.0, 0.0))
            println(sqrt(hypot(3.0, 4.0) * hypot(3.0, 4.0)))
        }
    "#);
    assert_eq!(out, &["5", "0", "5"]);
}
