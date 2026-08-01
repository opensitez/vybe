use super::helpers::run_prints;

// ── Rounding functions ───────────────────────────────────────────
#[test]
fn floor_ceil_round() {
    assert_eq!(
        run_prints(
            r#"<?php
echo floor(4.7);
echo floor(-4.3);
echo ceil(4.3);
echo ceil(-4.7);
echo round(4.5);
echo round(4.4);
echo round(-4.5);
"#
        ),
        &["4-55-454-5"]
    );
}

#[test]
fn round_precision() {
    assert_eq!(
        run_prints(
            r#"<?php
echo round(3.14159, 2);
echo round(3.14159, 4);
echo round(1234.5, -2);
"#
        ),
        &["3.143.14161200"]
    );
}

// ── Power / sqrt / log ───────────────────────────────────────────
#[test]
fn pow_sqrt() {
    assert_eq!(
        run_prints(
            r#"<?php
echo pow(2, 10);
echo sqrt(144);
echo sqrt(2);
"#
        ),
        // `echo` emits no newline, so PHP concatenates these into one line.
        &["1024121.4142135623731"]
    );
}

#[test]
fn log_functions() {
    assert_eq!(
        run_prints(
            r#"<?php
echo round(log(M_E), 2);
echo log10(1000);
echo log2(8);
"#
        ),
        &["133"]
    );
}

#[test]
fn exp_function() {
    assert_eq!(
        run_prints(
            r#"<?php
echo round(exp(0), 2);
echo round(exp(1), 4);
"#
        ),
        &["12.7183"]
    );
}

// ── min / max ────────────────────────────────────────────────────
#[test]
fn min_max_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo min(3, 1, 4, 1, 5, 9);
echo max(3, 1, 4, 1, 5, 9);
echo min([10, 20, 5, 30]);
echo max([10, 20, 5, 30]);
"#
        ),
        &["19530"]
    );
}

// ── fmod / intdiv ────────────────────────────────────────────────
#[test]
fn fmod_intdiv() {
    assert_eq!(
        run_prints(
            r#"<?php
echo fmod(10.5, 3);
echo intdiv(7, 2);
echo intdiv(10, 3);
"#
        ),
        &["1.533"]
    );
}

// ── Base conversion ──────────────────────────────────────────────
#[test]
fn base_convert_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo base_convert("ff", 16, 10);
echo base_convert("255", 10, 16);
echo base_convert("111", 2, 10);
echo base_convert("10", 10, 2);
"#
        ),
        &["255ff71010"]
    );
}

#[test]
fn bindec_decoct_dechex() {
    assert_eq!(
        run_prints(
            r#"<?php
echo bindec("1010");
echo decbin(42);
echo octdec("77");
echo decoct(255);
echo hexdec("ff");
echo dechex(255);
"#
        ),
        &["1010101063377255ff"]
    );
}

// ── Trigonometry ─────────────────────────────────────────────────
#[test]
fn trig_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo round(sin(0), 2);
echo round(cos(0), 2);
echo round(tan(0), 2);
echo round(sin(M_PI / 2), 2);
echo round(cos(M_PI), 2);
"#
        ),
        &["0101-1"]
    );
}

#[test]
fn trig_inverse() {
    assert_eq!(
        run_prints(
            r#"<?php
echo round(asin(1), 4);
echo round(acos(0), 4);
echo round(atan(1), 4);
echo round(atan2(1, 1), 4);
"#
        ),
        &["1.57081.57080.78540.7854"]
    );
}

// ── Hyperbolic functions ─────────────────────────────────────────
#[test]
fn hyperbolic_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo round(sinh(1), 4);
echo round(cosh(1), 4);
echo round(tanh(1), 4);
"#
        ),
        &["1.17521.54310.7616"]
    );
}

// ── Constants ────────────────────────────────────────────────────
#[test]
fn math_constants() {
    assert_eq!(
        run_prints(
            r#"<?php
echo round(M_PI, 5);
echo round(M_E, 5);
echo round(M_SQRT2, 5);
echo round(M_LN2, 5);
echo round(M_LOG2E, 5);
"#
        ),
        &["3.141592.718281.414210.693151.4427"]
    );
}

// ── Practical math ───────────────────────────────────────────────
#[test]
fn hypotenuse() {
    assert_eq!(
        run_prints(
            r#"<?php
echo hypot(3, 4);
echo hypot(5, 12);
"#
        ),
        &["513"]
    );
}

#[test]
fn clamp_pattern() {
    assert_eq!(
        run_prints(
            r#"<?php
function clamp(float $val, float $min, float $max): float {
    return max($min, min($max, $val));
}
echo clamp(15, 0, 10);
echo clamp(-5, 0, 10);
echo clamp(5, 0, 10);
"#
        ),
        &["1005"]
    );
}

#[test]
fn degrees_radians() {
    assert_eq!(
        run_prints(
            r#"<?php
echo round(deg2rad(180), 4);
echo round(rad2deg(M_PI), 0);
echo round(deg2rad(90), 4);
"#
        ),
        &["3.14161801.5708"]
    );
}

#[test]
fn is_nan_is_infinite() {
    assert_eq!(
        run_prints(
            r#"<?php
echo is_nan(sqrt(-1)) ? "nan" : "number";
echo is_infinite(log(0)) ? "inf" : "finite";
echo is_finite(42) ? "finite" : "not";
"#
        ),
        &["naninffinite"]
    );
}

#[test]
fn exponentiation_operator_negative_exponent() {
    assert_eq!(
        run_prints(
            r#"<?php
echo 2 ** 3;
echo "\n";
echo 2 ** 0;
echo "\n";
echo 2 ** -1;
echo "\n";
echo 9 ** 0.5;
"#
        ),
        vec!["8", "1", "0.5", "3"]
    );
}

#[test]
fn pow_and_exponentiation_precedence() {
    assert_eq!(
        run_prints(
            r#"<?php
echo 2 + 3 ** 2;
echo "\n";
echo (2 + 3) ** 2;
"#
        ),
        vec!["11", "25"]
    );
}

#[test]
fn fmod_sign_behavior() {
    assert_eq!(
        run_prints(
            r#"<?php
echo fmod(8.5, 2.5) . "\n";
echo fmod(-8.5, 2.5) . "\n";
echo fmod(8.5, -2.5) . "\n";
echo fmod(-8.5, -2.5) . "\n";
"#
        ),
        vec!["1", "-1", "1", "-1"]
    );
}

#[test]
fn float_division_quotients() {
    assert_eq!(
        run_prints(
            r#"<?php
echo (5 / 2) . "\n";
echo (5.0 / 2) . "\n";
echo (-5 / 2) . "\n";
echo (5 / -2) . "\n";
"#
        ),
        vec!["2.5", "2.5", "-2.5", "-2.5"]
    );
}

#[test]
fn modf_integer_and_fraction() {
    assert_eq!(
        run_prints(
            r#"<?php
[$frac, $int] = modf(3.25);
echo $frac . "\n";
echo $int . "\n";
[$frac2, $int2] = modf(-7.75);
echo $frac2 . "\n";
echo $int2 . "\n";
"#
        ),
        vec!["0.25", "3", "-0.75", "-7"]
    );
}

#[test]
fn math_constants_reference() {
    assert_eq!(
        run_prints(
            r#"<?php
echo (round(M_2_PI, 8) == 0.63661977 ? 'ok' : 'bad') . "\n";
echo (round(M_SQRT3, 8) == 1.73205081 ? 'ok' : 'bad') . "\n";
echo (M_LOG10E > 0 && M_LN10 > 0 ? 'ok' : 'bad') . "\n";
echo (M_LN2 < M_LN10 ? 'ok' : 'bad') . "\n";
"#
        ),
        vec!["ok", "ok", "ok", "ok"]
    );
}

#[test]
fn degree_radian_roundtrip_precision() {
    assert_eq!(
        run_prints(
            r#"<?php
echo round(rad2deg(deg2rad(45)), 12) . "\n";
echo round(deg2rad(rad2deg(M_PI)), 12) . "\n";
"#
        ),
        vec!["45", "3.14159265359"]
    );
}

#[test]
fn number_format_currency_style() {
    assert_eq!(
        run_prints(
            r#"<?php
echo number_format(1987.65, 1, ',', '.') . "\n";
echo number_format(1987.65, 3, '.', ',') . "\n";
echo number_format(1000, 0, '.', ',') . "\n";
"#
        ),
        // `1.987,7`: php rounds half-away-from-zero, so 1987.65 at 1 decimal is
        // 1987.7. `1.987,6` was JS `toFixed` banker's rounding.
        vec!["1.987,7", "1,987.650", "1,000"]
    );
}
