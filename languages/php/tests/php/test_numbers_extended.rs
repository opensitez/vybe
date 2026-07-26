use super::helpers::run_prints;

// ── Integer operations ────────────────────────────────────────

#[test]
fn integer_division_truncates() {
    assert_eq!(run_prints(r#"<?php echo intdiv(7, 2); "#), vec!["3"]);
}
#[test]
fn modulo_sign_follows_dividend() {
    assert_eq!(run_prints(r#"<?php echo -7 % 3; "#), vec!["-1"]);
}
#[test]
fn exponentiation_operator() {
    assert_eq!(run_prints(r#"<?php echo 2 ** 10; "#), vec!["1024"]);
}
#[test]
fn modulo_with_negative_divisor() {
    assert_eq!(run_prints(r#"<?php echo 7 % -3; "#), vec!["1"]);
}
#[test]
fn large_number_precision() {
    assert_eq!(
        run_prints(r#"<?php echo PHP_INT_MAX + 1 > PHP_INT_MAX ? 'overflow' : 'same'; "#),
        vec!["same"]
    );
}
#[test]
fn integer_cast_from_boolean_and_string() {
    assert_eq!(
        run_prints(
            r#"<?php
echo (int) true;
echo '|';
echo (int) false;
echo '|';
echo (int) "42";
echo '|';
echo (int) "04";
"#
        ),
        vec!["1|0|42|4"]
    );
}

// ── Float operations ──────────────────────────────────────────

#[test]
fn float_precision_comparison() {
    assert_eq!(
        run_prints(r#"<?php echo abs(0.1 + 0.2 - 0.3) < PHP_FLOAT_EPSILON ? 'equal' : 'diff'; "#),
        vec!["equal"]
    );
}
#[test]
fn float_nan_properties() {
    assert_eq!(
        run_prints(
            r#"<?php $n = NAN; echo is_nan($n) ? 'nan' : 'not'; echo ($n === $n) ? 'eq' : 'neq'; "#
        ),
        vec!["nanneq"]
    );
}
#[test]
fn float_inf_operations() {
    assert_eq!(
        run_prints(
            r#"<?php echo INF + 1 === INF ? 'same' : 'diff'; echo -INF < 0 ? 'neg' : 'pos'; "#
        ),
        vec!["sameneg"]
    );
}
#[test]
fn scientific_notation_float() {
    assert_eq!(run_prints(r#"<?php echo 1.5e3; "#), vec!["1500"]);
}
#[test]
fn float_modulus_and_rounding_sign() {
    assert_eq!(
        run_prints(r#"<?php echo fmod(-5.3, 2.0) . '|' . round(-2.5); "#),
        vec!["-1.3|-2"]
    );
}

#[test]
fn power_precedence_and_associativity() {
    assert_eq!(
        run_prints(r#"<?php
echo 2 ** 3 ** 2;
echo '|';
echo (2 ** 3) ** 2;
"#),
        vec!["512|64"]
    );
}

#[test]
fn unary_minus_vs_power_precedence() {
    assert_eq!(
        run_prints(r#"<?php
echo -2 ** 3;
echo '|';
echo (-2) ** 3;
echo '|';
echo -(2 ** 3);
"#),
        vec!["-8|-8|-8"]
    );
}

#[test]
fn modulo_precedence_vs_addition_and_subtraction() {
    assert_eq!(
        run_prints(r#"<?php
echo 10 + 9 % 3;
echo '|';
echo 10 - 9 % 3;
echo '|';
echo (10 + 9) % 3;
"#),
        vec!["10|10|1"]
    );
}

#[test]
    fn number_truthiness_runtime_checks() {
        assert_eq!(
            run_prints(r#"<?php
echo (0 == false) ? 'T' : 'F';
echo (0 === false) ? 'T' : 'F';
echo (0.0 == false) ? 'T' : 'F';
echo ('' == false) ? 'T' : 'F';
echo ('0' == false) ? 'T' : 'F';
"#),
            vec!["TFTTT"]
    );
}

// ── Number functions ──────────────────────────────────────────

#[test]
fn abs_of_various() {
    assert_eq!(
        run_prints(r#"<?php echo abs(-5) . ',' . abs(5) . ',' . abs(-3.14); "#),
        vec!["5,5,3.14"]
    );
}
#[test]
fn max_min_with_various_types() {
    assert_eq!(
        run_prints(r#"<?php echo max('10', 9, 11) . ',' . min(0, -1, 1); "#),
        vec!["11,-1"]
    );
}
#[test]
fn round_various_precisions() {
    assert_eq!(
        run_prints(r#"<?php echo round(1.555, 2) . ',' . round(1234.5, -2); "#),
        vec!["1.56,1200"]
    );
}
#[test]
fn number_format_custom_decimals() {
    assert_eq!(
        run_prints(
            r#"<?php echo number_format(1234.56789, 2, ',', ' '); echo '|'; echo number_format(12, 0, '.', ','); "#),
        vec!["1234,57|12"]
    );
}

// ── Random number generation ──────────────────────────────────

#[test]
fn rand_range() {
    assert_eq!(
        run_prints(
            r#"<?php
$n = rand(1, 10);
echo ($n >= 1 && $n <= 10) ? 'in_range' : 'out';
"#
        ),
        vec!["in_range"]
    );
}
#[test]
fn mt_rand_consistent_range() {
    assert_eq!(
        run_prints(
            r#"<?php
$n = mt_rand(100, 200);
echo ($n >= 100 && $n <= 200) ? 'ok' : 'fail';
"#
        ),
        vec!["ok"]
    );
}
#[test]
fn random_int_range() {
    assert_eq!(
        run_prints(
            r#"<?php
$n = random_int(1, 6);
echo ($n >= 1 && $n <= 6) ? 'ok' : 'fail';
"#
        ),
        vec!["ok"]
    );
}

// ── Number formatting ─────────────────────────────────────────

#[test]
fn number_format_zero_decimals() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1234567); "#),
        vec!["1,234,567"]
    );
}
#[test]
fn decoct_decbin_dechex() {
    assert_eq!(
        run_prints(r#"<?php echo decoct(8) . ',' . decbin(10) . ',' . dechex(255); "#),
        vec!["10,1010,ff"]
    );
}
#[test]
fn octal_or_hex_conversion_from_prefixes() {
    assert_eq!(
        run_prints(r#"<?php echo octdec('075'); echo '|'; echo hexdec('0x1f'); "#),
        vec!["61|31"]
    );
}
#[test]
fn octdec_bindec_hexdec() {
    assert_eq!(
        run_prints(r#"<?php echo octdec('10') . ',' . bindec('1010') . ',' . hexdec('ff'); "#),
        vec!["8,10,255"]
    );
}

// ── GMP (arbitrary precision) ────────────────────────────────

#[test]
fn bcmath_add() {
    assert_eq!(
        run_prints(r#"<?php echo bcadd('999999999999999999', '1'); "#),
        vec!["1000000000000000000"]
    );
}
#[test]
fn bcmath_mul() {
    assert_eq!(
        run_prints(r#"<?php echo bcmul('123456789', '987654321'); "#),
        vec!["121932631112635269"]
    );
}
#[test]
fn bcmath_div() {
    assert_eq!(
        run_prints(r#"<?php echo bcdiv('10', '3', 4); "#),
        vec!["3.3333"]
    );
}
#[test]
fn bcmath_pow() {
    assert_eq!(
        run_prints(r#"<?php echo bcpow('2', '64'); "#),
        vec!["18446744073709551616"]
    );
}
#[test]
fn bcmath_comp() {
    assert_eq!(
        run_prints(r#"<?php echo bccomp('1.23456789', '1.23456790'); "#),
        vec!["0"]
    );
}
#[test]
fn bcmath_scale() {
    assert_eq!(
        run_prints(r#"<?php bcscale(2); echo bcadd('1', '2'); "#),
        vec!["3.00"]
    );
}

#[test]
fn float_cast_roundtrip_with_sign_and_fractional_input() {
    assert_eq!(
        run_prints(
            r#"<?php
echo (int) 3.99;
echo '|';
echo (int) -3.99;
echo '|';
echo (float) '12.34';
"#
        ),
        vec!["3|-3|12.34"]
    );
}

#[test]
fn division_and_precedence_edges() {
    assert_eq!(
        run_prints(
            r#"<?php
echo 10 / 4;
echo '|';
echo intdiv(10, 4);
echo '|';
echo (1 + 2) / 3;
echo '|';
echo 1 + 2 / 3;
            "#
        ),
        vec!["2.5|2|1|1.6666666666667"]
    );
}

#[test]
fn modulo_and_negative_zero_like_edges() {
    assert_eq!(
        run_prints(
            r#"<?php
echo 0 % 3;
echo '|';
echo 5 % -2;
echo '|';
echo -5 % 2;
echo '|';
echo (-5) % (-2);
echo '|';
echo (-0) == 0;
"#
        ),
        vec!["0|1|-1|-1|1"]
    );
}

#[test]
fn comparison_with_nan_and_infinity() {
    assert_eq!(
        run_prints(
            r#"<?php
echo is_nan(NAN) ? 'nan' : 'no';
echo '|';
echo (INF > 1e308) ? 'big' : 'sm';
echo '|';
echo is_infinite(-INF) ? 'inf' : 'no';
echo '|';
echo (NAN === NAN) ? 'eq' : 'neq';
"#
        ),
        vec!["nan|big|inf|neq"]
    );
}

#[test]
fn scientific_and_signed_notation_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo 1.2e2;
echo '|';
echo 1.2e-2;
echo '|';
echo -1.2e1;
echo '|';
echo 1.0e0;
            "#
        ),
        vec!["120|0.012|-12|1"]
    );
}

#[test]
fn number_base_parsing_and_formatting_edges() {
    assert_eq!(
        run_prints(
            r#"<?php
echo 0b1_0_1_0;
echo '|';
echo 0o10;
echo '|';
echo 0x1a;
echo '|';
echo base_convert('ff', 16, 2);
echo '|';
echo bindec('00001111');
"#
        ),
        vec!["10|8|26|11111111|15"]
    );
}
