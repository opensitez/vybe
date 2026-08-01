use crate::helpers::run_prints;

#[test]
fn test_abs_zero_and_positive_integers() {
    let out = run_prints(
        r#"
        fun main() {
            println(abs(0))
            println(abs(12))
            println(abs(999))
        }
    "#,
    );
    assert_eq!(out, &["0", "12", "999"]);
}

#[test]
fn test_abs_negative_integer_and_long() {
    let out = run_prints(
        r#"
        fun main() {
            println(abs(-12))
            println(abs(-12345))
            println(abs(-123L))
            println(abs(-999_999_999L))
        }
    "#,
    );
    assert_eq!(out, &["12", "12345", "123", "999999999"]);
}

#[test]
fn test_abs_double_negative_and_zero() {
    let out = run_prints(
        r#"
        fun main() {
            println(abs(-12.75))
            println(abs(0.0))
            println(abs(-0.0))
            println(abs(5.0))
        }
    "#,
    );
    assert_eq!(out, &["12.75", "0", "0", "5"]);
}

#[test]
fn test_max_basic() {
    let out = run_prints(
        r#"
        fun main() {
            println(max(3, 7))
            println(max(-5, 2))
            println(max(9, 9))
        }
    "#,
    );
    assert_eq!(out, &["7", "2", "9"]);
}

#[test]
fn test_min_basic() {
    let out = run_prints(
        r#"
        fun main() {
            println(min(3, 7))
            println(min(-5, 2))
            println(min(9, 9))
        }
    "#,
    );
    assert_eq!(out, &["3", "-5", "9"]);
}

#[test]
fn test_max_is_commutative_for_negative_values() {
    let out = run_prints(
        r#"
        fun main() {
            println(max(-3, -10))
            println(max(-10, -3))
            println(min(-3, -10))
            println(min(-10, -3))
        }
    "#,
    );
    assert_eq!(out, &["-3", "-3", "-10", "-10"]);
}

#[test]
fn test_min_max_chain() {
    let out = run_prints(
        r#"
        fun main() {
            println(max(1, max(3, 9)))
            println(min(1, min(3, 9)))
            println(max(min(10, 2), max(-4, 12)))
            println(min(max(10, 2), min(-4, 12)))
        }
    "#,
    );
    assert_eq!(out, &["9", "1", "12", "-4"]);
}

#[test]
fn test_max_with_doubles() {
    let out = run_prints(
        r#"
        fun main() {
            println(max(1.5, 2.5))
            println(min(1.5, -2.5))
            println(min(-1.2, -3.4))
            println(max(-1.2, -3.4))
        }
    "#,
    );
    assert_eq!(out, &["2.5", "-2.5", "-3.4", "-1.2"]);
}

#[test]
fn test_pow_positive_exponent_integers_as_doubles() {
    let out = run_prints(
        r#"
        fun main() {
            println(pow(2.0, 0.0))
            println(pow(2.0, 1.0))
            println(pow(2.0, 2.0))
            println(pow(2.0, 3.0))
            println(pow(2.0, 4.0))
        }
    "#,
    );
    assert_eq!(out, &["1", "1", "4", "8", "16"]);
}

#[test]
fn test_pow_fractional_exponent_small_roots() {
    let out = run_prints(
        r#"
        fun main() {
            println(pow(27.0, 1.0 / 3.0))
            println(pow(64.0, 1.0 / 3.0))
            println(pow(4.0, 0.5))
            println(pow(16.0, 0.5))
            println(round(pow(27.0, 1.0 / 3.0)))
        }
    "#,
    );
    assert_eq!(out, &["3", "4", "2", "4", "3"]);
}

#[test]
fn test_pow_negative_base_even_and_odd_exponents() {
    let out = run_prints(
        r#"
        fun main() {
            println(pow(-2.0, 2.0))
            println(pow(-2.0, 3.0))
            println(pow(-3.0, 4.0))
            println(pow(-3.0, 5.0))
        }
    "#,
    );
    assert_eq!(out, &["4", "-8", "81", "-243"]);
}

#[test]
fn test_pow_negative_exponent() {
    let out = run_prints(
        r#"
        fun main() {
            println(pow(2.0, -1.0))
            println(pow(4.0, -2.0))
            println(pow(10.0, -1.0))
        }
    "#,
    );
    assert_eq!(out, &["0.5", "0.0625", "0.1"]);
}

#[test]
fn test_pow_zero_base_cases() {
    let out = run_prints(
        r#"
        fun main() {
            println(pow(0.0, 5.0))
            println(pow(0.0, 0.0))
            println(pow(0.0, 1.0))
        }
    "#,
    );
    assert_eq!(out, &["0", "1", "0"]);
}

#[test]
fn test_sqrt_identity_and_zero() {
    let out = run_prints(
        r#"
        fun main() {
            println(sqrt(0.0))
            println(sqrt(1.0))
            println(sqrt(4.0))
            println(sqrt(81.0))
        }
    "#,
    );
    assert_eq!(out, &["0", "1", "2", "9"]);
}

#[test]
fn test_sqrt_roundtrip_square_values() {
    let out = run_prints(
        r#"
        fun main() {
            println(sqrt(144.0))
            println(sqrt(225.0))
            println(sqrt(2.0 * 2.0))
        }
    "#,
    );
    assert_eq!(out, &["12", "15", "2"]);
}

#[test]
fn test_sqrt_decimal_input() {
    let out = run_prints(
        r#"
        fun main() {
            println(sqrt(12.25))
            println(sqrt(0.81))
            println(sqrt(6.25))
            println(sqrt(2.56))
        }
    "#,
    );
    assert_eq!(out, &["3.5", "0.9", "2.5", "1.6"]);
}

#[test]
fn test_sqrt_negative_is_not_a_number() {
    let out = run_prints(
        r#"
        fun main() {
            println(sqrt(-4.0).isNaN())
            println(sqrt(-1.5).isNaN())
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_floor_for_positive_and_zero_values() {
    let out = run_prints(
        r#"
        fun main() {
            println(floor(3.0))
            println(floor(3.9))
            println(floor(0.1))
            println(floor(-0.0))
        }
    "#,
    );
    assert_eq!(out, &["3", "3", "0", "0"]);
}

#[test]
fn test_floor_negative_values() {
    let out = run_prints(
        r#"
        fun main() {
            println(floor(-3.2))
            println(floor(-3.9))
            println(floor(-0.1))
            println(floor(-2.0))
        }
    "#,
    );
    assert_eq!(out, &["-4", "-4", "-1", "-2"]);
}

#[test]
fn test_ceil_for_positive_and_zero_values() {
    let out = run_prints(
        r#"
        fun main() {
            println(ceil(3.0))
            println(ceil(3.2))
            println(ceil(0.0))
            println(ceil(0.9))
        }
    "#,
    );
    assert_eq!(out, &["3", "4", "0", "1"]);
}

#[test]
fn test_ceil_negative_values() {
    let out = run_prints(
        r#"
        fun main() {
            println(ceil(-3.2))
            println(ceil(-3.9))
            println(ceil(-0.9))
            println(ceil(-2.0))
        }
    "#,
    );
    assert_eq!(out, &["-3", "-3", "0", "-2"]);
}

#[test]
fn test_round_half_and_fractional() {
    let out = run_prints(
        r#"
        fun main() {
            println(round(2.4))
            println(round(2.5))
            println(round(2.6))
            println(round(-2.4))
            println(round(-2.5))
            println(round(-2.6))
        }
    "#,
    );
    assert_eq!(out, &["2", "3", "3", "-2", "-3", "-3"]);
}

#[test]
fn test_round_small_integers() {
    let out = run_prints(
        r#"
        fun main() {
            println(round(0.0))
            println(round(-0.0))
            println(round(0.49))
            println(round(-0.49))
            println(round(999.6))
            println(round(-999.6))
        }
    "#,
    );
    assert_eq!(out, &["0", "0", "0", "0", "1000", "-1000"]);
}

#[test]
fn test_trig_zero_identity() {
    let out = run_prints(
        r#"
        fun main() {
            println(sin(0.0))
            println(cos(0.0))
            println(tan(0.0))
        }
    "#,
    );
    assert_eq!(out, &["0", "1", "0"]);
}

#[test]
fn test_abs_and_round_combo() {
    let out = run_prints(
        r#"
        fun main() {
            println(abs(round(-2.6) + pow(2.0, 3.0)))
            println(abs(round(3.4) - pow(3.0, 2.0)))
            println(abs(pow(2.0, 2.0) - pow(2.0, 3.0) + round(0.5)))
        }
    "#,
    );
    assert_eq!(out, &["11", "1", "3"]);
}

#[test]
fn test_math_pipeline_with_floor_and_round() {
    let out = run_prints(
        r#"
        fun main() {
            val raw = abs(floor(-3.7) + ceil(3.2))
            val rounded = round(raw / 2.0)
            println(raw)
            println(rounded)
        }
    "#,
    );
    assert_eq!(out, &["7", "4"]);
}

#[test]
fn test_math_round_trip_for_integer_doubles() {
    let out = run_prints(
        r#"
        fun main() {
            val a = floor(5.9)
            val b = ceil(5.1)
            val c = round(5.2)
            println(a)
            println(b)
            println(c)
        }
    "#,
    );
    assert_eq!(out, &["5", "6", "5"]);
}

#[test]
fn test_math_composition_with_min_max() {
    let out = run_prints(
        r#"
        fun main() {
            val score = max(abs(-12), 8)
            val margin = min(4.7, 9.2)
            println(score)
            println(margin)
            println(score + margin)
        }
    "#,
    );
    assert_eq!(out, &["12", "4.7", "16.7"]);
}

#[test]
fn test_pow_small_fractional_exponent() {
    let out = run_prints(
        r#"
        fun main() {
            println(pow(16.0, 0.5))
            println(pow(16.0, 0.25))
            println(pow(1.0, 999.0))
            println(pow(9.0, 0.5))
        }
    "#,
    );
    assert_eq!(out, &["4", "2", "1", "3"]);
}

#[test]
fn test_abs_plus_pow_and_rounding_stability() {
    let out = run_prints(
        r#"
        fun main() {
            val x = abs(pow(-2.0, 2.0) - 3.0)
            val y = round(2.5)
            val z = floor(5.9 - 1.2)
            println(x)
            println(y)
            println(z)
            println(x + y + z)
        }
    "#,
    );
    assert_eq!(out, &["1", "3", "4", "8"]);
}

#[test]
fn test_log_exp_roundtrip_for_e_identity() {
    let out = run_prints(
        r#"
        fun main() {
            val eRounded = kotlin.math.round(kotlin.math.exp(1.0) * 1000.0) / 1000.0
            val recovered = kotlin.math.ln(kotlin.math.exp(1.0))
            println(recovered)
            println(eRounded > 2.7)
        }
    "#,
    );
    assert_eq!(out, &["1.0", "true"]);
}

#[test]
fn test_log_base_arithmetic_and_zero_boundary() {
    let out = run_prints(
        r#"
        fun main() {
            val ten = kotlin.math.log(1000.0, 10.0)
            val two = kotlin.math.log(8.0, 2.0)
            val tiny = kotlin.math.log10(1.0)
            println(kotlin.math.round(ten))
            println(kotlin.math.round(two))
            println(tiny)
        }
    "#,
    );
    assert_eq!(out, &["3", "3", "0.0"]);
}

#[test]
fn test_hypot_uses_pythagorean_contract() {
    let out = run_prints(
        r#"
        fun main() {
            val h = kotlin.math.hypot(3.0, 4.0)
            println(h)
            println(hypot(5.0, 12.0))
        }
    "#,
    );
    assert_eq!(out, &["5.0", "13.0"]);
}

#[test]
fn test_trig_inverse_relationships() {
    let out = run_prints(
        r#"
        fun main() {
            val angle = kotlin.math.asin(kotlin.math.sin(1.0))
            val cosv = kotlin.math.cos(angle)
            val diff = kotlin.math.round((kotlin.math.abs(cosv - kotlin.math.cos(1.0)) * 1e6).toDouble())
            println(kotlin.math.abs(angle) < 1e-9)
            println(diff == 0.0)
        }
    "#,
    );
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_atan2_quadrant_and_zero_axis_signals() {
    let out = run_prints(
        r#"
        fun main() {
            val a = kotlin.math.atan2(0.0, -1.0)
            val b = kotlin.math.atan2(1.0, 0.0)
            println(a == kotlin.math.PI)
            println(b > 1.5)
            println(b < 2.0)
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_floor_div_of_ints_and_signals() {
    let out = run_prints(
        r#"
        fun main() {
            println((-5).floorDiv(2))
            println((-5).mod(2))
            println(5.floorDiv(-2))
            println(5.mod(-2))
        }
    "#,
    );
    assert_eq!(out, &["-3", "-1", "-3", "1"]);
}

#[test]
fn test_unsigned_power_and_sqrt_chain_edge_case() {
    let out = run_prints(
        r#"
        fun main() {
            val squared = kotlin.math.sqrt(2.0) * kotlin.math.sqrt(2.0)
            val closeToTwo = kotlin.math.abs(squared - 2.0) < 0.0000001
            println(closeToTwo)
            println(kotlin.math.pow(2.0, 0.0))
            println(kotlin.math.pow(0.0, 1.0))
            println(kotlin.math.pow(0.0, 0.0))
        }
    "#,
    );
    assert_eq!(out, &["true", "1.0", "0.0", "1.0"]);
}

#[test]
fn test_math_next_toward_positive_infinity() {
    let out = run_prints(
        r#"
        fun main() {
            val start = 1.0
            val next = kotlin.math.nextUp(start)
            val moved = next > start
            println(moved)
            val down = kotlin.math.nextDown(next)
            println(down <= next)
            println(start == kotlin.math.nextDown(next))
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_nan_and_infinity_propagation() {
    let out = run_prints(
        r#"
        fun main() {
            val nan = kotlin.math.sqrt(-1.0)
            val inf = 1.0 / 0.0
            val finite = kotlin.math.isFinite(inf)
            val notFinite = kotlin.math.isFinite(nan)
            println(nan.isNaN())
            println(finite)
            println(notFinite)
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "false"]);
}

#[test]
fn test_ulp_precision_invariant() {
    let out = run_prints(
        r#"
        fun main() {
            val v = 1.0
            val step = kotlin.math.ulp(v)
            val nearOne = 1.0 + step
            val isAdjacent = kotlin.math.nextAfter(1.0, Double.POSITIVE_INFINITY) == nearOne
            println(nearOne > v)
            println(step > 0.0)
            println(isAdjacent)
            println(kotlin.math.abs(step - kotlin.math.ulp(nearOne)) < 1e-20)
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true", "true"]);
}
