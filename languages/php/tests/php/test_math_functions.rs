use super::helpers::{compile_ok, run_prints};

// ── Trigonometry ──────────────────────────────────────────────

#[test]
fn sin_cos_pi() {
    assert_eq!(
        run_prints(r#"<?php echo round(sin(M_PI / 2), 10) . ',' . round(cos(M_PI), 10); "#),
        vec!["1,-1"]
    );
}
#[test]
fn tan_45_degrees() {
    assert_eq!(
        run_prints(r#"<?php echo round(tan(deg2rad(45)), 10); "#),
        vec!["1"]
    );
}
#[test]
fn atan2_quadrant() {
    assert_eq!(
        run_prints(r#"<?php echo round(atan2(1, 1), 5); "#),
        vec!["0.7854"]
    );
}
#[test]
fn degrees_radians_conversion() {
    assert_eq!(
        run_prints(r#"<?php echo round(rad2deg(M_PI), 5); "#),
        vec!["180"]
    );
}
#[test]
fn asin_acos_atan_identity() {
    assert_eq!(
        run_prints(r#"<?php echo round(asin(sin(0.5)), 10); "#),
        vec!["0.5"]
    );
}
#[test]
fn hypot_pythagorean() {
    assert_eq!(run_prints(r#"<?php echo hypot(3, 4); "#), vec!["5"]);
}

// ── Logarithm and exponential ─────────────────────────────────

#[test]
fn log_natural() {
    assert_eq!(run_prints(r#"<?php echo round(log(M_E), 10); "#), vec!["1"]);
}
#[test]
fn log_base_10() {
    assert_eq!(run_prints(r#"<?php echo log10(100); "#), vec!["2"]);
}
#[test]
fn log_base_2() {
    assert_eq!(run_prints(r#"<?php echo log(8, 2); "#), vec!["3"]);
}
#[test]
fn exp_e_squared() {
    assert_eq!(
        run_prints(r#"<?php echo round(exp(2), 5); "#),
        vec!["7.38906"]
    );
}
#[test]
fn sqrt_perfect_square() {
    assert_eq!(run_prints(r#"<?php echo sqrt(144); "#), vec!["12"]);
}
#[test]
fn pow_function() {
    assert_eq!(run_prints(r#"<?php echo pow(2, 10); "#), vec!["1024"]);
}

// ── abs / sign ────────────────────────────────────────────────

#[test]
fn abs_negative_int() {
    assert_eq!(run_prints(r#"<?php echo abs(-42); "#), vec!["42"]);
}
#[test]
fn abs_negative_float() {
    assert_eq!(run_prints(r#"<?php echo abs(-3.14); "#), vec!["3.14"]);
}

// ── min / max ─────────────────────────────────────────────────

#[test]
fn min_multiple_args() {
    assert_eq!(
        run_prints(r#"<?php echo min(3, 1, 4, 1, 5, 9); "#),
        vec!["1"]
    );
}
#[test]
fn max_multiple_args() {
    assert_eq!(
        run_prints(r#"<?php echo max(3, 1, 4, 1, 5, 9); "#),
        vec!["9"]
    );
}
#[test]
fn min_array_arg() {
    assert_eq!(run_prints(r#"<?php echo min([5, 2, 8, 1]); "#), vec!["1"]);
}
#[test]
fn max_array_arg() {
    assert_eq!(run_prints(r#"<?php echo max([5, 2, 8, 1]); "#), vec!["8"]);
}

// ── base_convert / number bases ───────────────────────────────

#[test]
fn base_convert_hex_to_decimal() {
    assert_eq!(
        run_prints(r#"<?php echo base_convert('ff', 16, 10); "#),
        vec!["255"]
    );
}
#[test]
fn base_convert_decimal_to_binary() {
    assert_eq!(
        run_prints(r#"<?php echo base_convert('255', 10, 2); "#),
        vec!["11111111"]
    );
}
#[test]
fn decbin_hexdec() {
    assert_eq!(
        run_prints(r#"<?php echo decbin(10) . ',' . hexdec('1A'); "#),
        vec!["1010,26"]
    );
}

// ── Number formatting ─────────────────────────────────────────

#[test]
fn number_format_integer() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1000000, 0, '.', ','); "#),
        vec!["1,000,000"]
    );
}
#[test]
fn printf_padding_and_precision() {
    assert_eq!(
        run_prints(r#"<?php printf('%08.2f', 3.14); "#),
        vec!["00003.14"]
    );
}

// ── PHP math constants ────────────────────────────────────────

#[test]
fn m_pi_constant() {
    assert_eq!(
        run_prints(r#"<?php echo round(M_PI, 5); "#),
        vec!["3.14159"]
    );
}
#[test]
fn m_e_constant() {
    assert_eq!(run_prints(r#"<?php echo round(M_E, 5); "#), vec!["2.71828"]);
}
#[test]
fn m_sqrt2_constant() {
    assert_eq!(
        run_prints(r#"<?php echo round(M_SQRT2, 5); "#),
        vec!["1.41421"]
    );
}
#[test]
fn m_ln2_constant() {
    assert_eq!(
        run_prints(r#"<?php echo round(M_LN2, 5); "#),
        vec!["0.69315"]
    );
}

#[test]
fn math_floor_and_ceil_behavior_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo floor(4.7);
echo '|';
echo ceil(4.1);
"#,
        ),
        vec!["4|5"]
    );
}

#[test]
fn math_round_half_even_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo round(2.5);
echo '|';
echo round(3.5, 0, PHP_ROUND_HALF_UP);
"#,
        ),
        vec!["2|4"]
    );
}

#[test]
fn math_modf_returns_fraction_and_integer_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
list($frac, $int) = modf(3.75);
echo $frac;
echo '|';
echo $int;
"#,
            ),
        vec!["0.75|3"]
    );
}

#[test]
fn math_radix_conversion_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo base_convert(255, 10, 16);
echo '|';
echo octdec('10');
"#,
        ),
        vec!["ff|8"]
    );
}

#[test]
fn intdiv_negative_rounding_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo intdiv(7, 3);
echo '|';
echo intdiv(-7, 3);
echo '|';
echo intdiv(7, -3);
"#,
        ),
        vec!["2|-2|-2"]
    );
}

#[test]
fn fdiv_zero_division_behavior_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo fdiv(5.0, 0.0);
echo '|';
echo fdiv(-5.0, 0.0);
echo '|';
echo is_infinite(fdiv(1, 0)) ? '1' : '0';
"#,
        ),
        vec!["INF|-INF|1"]
    );
}

#[test]
fn trig_periodic_relationship_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo round(sin(M_PI), 10);
echo '|';
echo round(cos(0), 10);
echo '|';
echo round(sin(M_PI / 6), 10);
"#,
        ),
        vec!["0|1|0.5"]
    );
}

#[test]
fn ceil_and_floor_negative_values_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo ceil(-3.2);
echo '|';
echo floor(-3.2);
"#,
        ),
        vec!["-3|-4"]
    );
}

#[test]
fn log_invalid_base_and_base_e_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo log(8, 2);
echo '|';
echo is_nan(log(-1)) ? '1' : '0';
echo '|';
echo round(exp(log(5)), 5);
"#,
        ),
        vec!["3|1|5"]
    );
}

#[test]
fn power_edge_cases_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo pow(0, 0);
echo '|';
echo pow(-2, 3);
echo '|';
echo pow(-2, 2);
"#,
        ),
        vec!["1|-8|4"]
    );
}

#[test]
fn number_format_negative_and_precision_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo number_format(-1234.5678, 2, ',', '.');
echo '|';
echo number_format(99.9, 0, ',', '');
"#,
        ),
        vec!["-1.234,57|100"]
    );
}

#[test]
fn decimal_and_binary_mix_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo decbin(5) . '|';
echo bindec('101');
"#,
        ),
        vec!["101|5"]
    );
}
