use super::helpers::run_prints;

// ── Integer limits and size ──────────────────────────────────────

#[test]
fn php_int_max_min_size() {
    assert_eq!(run_prints(r#"<?php
echo PHP_INT_SIZE . "\n";
echo (PHP_INT_MAX > 0) ? 'max_positive' : 'max_negative';
echo "\n";
echo (PHP_INT_MIN < 0) ? 'min_negative' : 'min_positive';
echo "\n";
"#), vec!["8","max_positive","min_negative"]);
}

#[test]
fn integer_overflow_to_float() {
    assert_eq!(run_prints(r#"<?php
$max = PHP_INT_MAX;
$overflow = $max + 1;
echo is_float($overflow) ? 'float' : 'int';
echo "\n";
echo is_int($max) ? 'int' : 'float';
echo "\n";
"#), vec!["float","int"]);
}

#[test]
fn php_float_epsilon_and_max() {
    assert_eq!(run_prints(r#"<?php
echo (PHP_FLOAT_EPSILON > 0) ? 'positive' : 'zero';
echo "\n";
echo (PHP_FLOAT_MAX > 1e100) ? 'large' : 'small';
echo "\n";
echo PHP_FLOAT_DIG . "\n";
"#), vec!["positive","large","15"]);
}

#[test]
fn inf_and_is_infinite() {
    assert_eq!(run_prints(r#"<?php
$inf = INF;
echo is_infinite($inf) ? 'infinite' : 'finite';
echo "\n";
echo is_infinite(-INF) ? 'infinite' : 'finite';
echo "\n";
echo is_infinite(1.0) ? 'infinite' : 'finite';
echo "\n";
echo $inf > PHP_INT_MAX ? 'bigger' : 'not bigger';
echo "\n";
"#), vec!["infinite","infinite","finite","bigger"]);
}

#[test]
fn nan_and_is_nan() {
    assert_eq!(run_prints(r#"<?php
$nan = NAN;
echo is_nan($nan) ? 'nan' : 'not nan';
echo "\n";
echo is_nan(0.0) ? 'nan' : 'not nan';
echo "\n";
echo ($nan === $nan) ? 'equal' : 'not equal';
echo "\n";
"#), vec!["nan","not nan","not equal"]);
}

#[test]
fn is_finite_check() {
    assert_eq!(run_prints(r#"<?php
echo is_finite(42.0) ? 'finite' : 'infinite';
echo "\n";
echo is_finite(INF) ? 'finite' : 'infinite';
echo "\n";
echo is_finite(NAN) ? 'finite' : 'infinite';
echo "\n";
echo is_finite(PHP_FLOAT_MAX) ? 'finite' : 'infinite';
echo "\n";
"#), vec!["finite","infinite","infinite","finite"]);
}

#[test]
fn intdiv_integer_division() {
    assert_eq!(run_prints(r#"<?php
echo intdiv(7, 2) . "\n";
echo intdiv(10, 3) . "\n";
echo intdiv(-7, 2) . "\n";
echo intdiv(100, 10) . "\n";
"#), vec!["3","3","-3","10"]);
}

#[test]
fn fmod_float_modulo() {
    assert_eq!(run_prints(r#"<?php
echo fmod(10.5, 3.0) . "\n";
echo fmod(10.0, 3.0) . "\n";
echo fmod(-10.5, 3.0) . "\n";
"#), vec!["1.5","1","-1.5"]);
}

#[test]
fn fdiv_division_inf_on_zero() {
    assert_eq!(run_prints(r#"<?php
echo fdiv(10, 0) === INF ? 'INF' : 'other';
echo "\n";
echo fdiv(-10, 0) === -INF ? '-INF' : 'other';
echo "\n";
echo fdiv(0, 0) !== fdiv(0, 0) ? 'NAN' : 'other';
echo "\n";
echo fdiv(10, 2) . "\n";
"#), vec!["INF","-INF","NAN","5"]);
}

#[test]
fn abs_on_negative_float() {
    assert_eq!(run_prints(r#"<?php
echo abs(-3.14) . "\n";
echo abs(3.14) . "\n";
echo abs(-0) . "\n";
echo abs(-PHP_INT_MAX) . "\n";
"#), vec!["3.14","3.14","0","9223372036854775807"]);
}

#[test]
fn ceil_floor_round_various() {
    assert_eq!(run_prints(r#"<?php
echo ceil(4.1) . "\n";
echo ceil(-4.1) . "\n";
echo floor(4.9) . "\n";
echo floor(-4.9) . "\n";
echo round(4.5) . "\n";
echo round(4.55, 1) . "\n";
echo round(-4.5) . "\n";
"#), vec!["5","-4","4","-5","5","4.6","-5"]);
}

#[test]
fn round_half_up_half_even() {
    assert_eq!(run_prints(r#"<?php
echo round(2.5, 0, PHP_ROUND_HALF_UP) . "\n";
echo round(3.5, 0, PHP_ROUND_HALF_UP) . "\n";
echo round(2.5, 0, PHP_ROUND_HALF_EVEN) . "\n";
echo round(3.5, 0, PHP_ROUND_HALF_EVEN) . "\n";
"#), vec!["3","4","2","4"]);
}

#[test]
fn number_format_separators() {
    assert_eq!(run_prints(r#"<?php
echo number_format(1234567.891, 2) . "\n";
echo number_format(1234567.891, 2, ',', '.') . "\n";
echo number_format(1000, 0, '.', '') . "\n";
"#), vec!["1,234,567.89","1.234.567,89","1000"]);
}

#[test]
fn hexadecimal_literal() {
    assert_eq!(run_prints(r#"<?php
$hex = 0xFF;
echo $hex . "\n";
echo gettype($hex) . "\n";
$neg = -0x0F;
echo $neg . "\n";
"#), vec!["255","integer","-15"]);
}

#[test]
fn octal_literal() {
    assert_eq!(run_prints(r#"<?php
$oct = 0777;
echo $oct . "\n";
$small = 010;
echo $small . "\n";
"#), vec!["511","8"]);
}

#[test]
fn binary_literal() {
    assert_eq!(run_prints(r#"<?php
$bin = 0b1010;
echo $bin . "\n";
$bin2 = 0b11111111;
echo $bin2 . "\n";
"#), vec!["10","255"]);
}

#[test]
fn underscore_numeric_literal() {
    assert_eq!(run_prints(r#"<?php
$million = 1_000_000;
echo $million . "\n";
$pi = 3.141_592_653;
echo round($pi, 6) . "\n";
$hex = 0xFF_FF;
echo $hex . "\n";
"#), vec!["1000000","3.141593","65535"]);
}

#[test]
fn bitwise_and_or_xor() {
    assert_eq!(run_prints(r#"<?php
echo (0b1100 & 0b1010) . "\n";
echo (0b1100 | 0b1010) . "\n";
echo (0b1100 ^ 0b1010) . "\n";
"#), vec!["8","14","6"]);
}

#[test]
fn bitwise_not() {
    assert_eq!(run_prints(r#"<?php
$a = 0b00001111;
echo (~$a & 0xFF) . "\n";
echo ~0 . "\n";
"#), vec!["240","-1"]);
}

#[test]
fn left_right_shift() {
    assert_eq!(run_prints(r#"<?php
echo (1 << 4) . "\n";
echo (256 >> 3) . "\n";
echo (0b0001 << 3) . "\n";
echo (0b10000 >> 2) . "\n";
"#), vec!["16","32","8","4"]);
}

#[test]
fn base_convert_between_bases() {
    assert_eq!(run_prints(r#"<?php
echo base_convert('ff', 16, 10) . "\n";
echo base_convert('255', 10, 16) . "\n";
echo base_convert('11111111', 2, 10) . "\n";
echo base_convert('377', 8, 10) . "\n";
"#), vec!["255","ff","255","255"]);
}

#[test]
fn bindec_decoct_dechex_decbin() {
    assert_eq!(run_prints(r#"<?php
echo bindec('1010') . "\n";
echo decbin(10) . "\n";
echo octdec('77') . "\n";
echo decoct(63) . "\n";
echo hexdec('ff') . "\n";
echo dechex(255) . "\n";
"#), vec!["10","1010","63","77","255","ff"]);
}

#[test]
fn pow_vs_exponentiation_operator() {
    assert_eq!(run_prints(r#"<?php
echo pow(2, 8) . "\n";
echo 2 ** 8 . "\n";
echo pow(3, 0) . "\n";
echo 2 ** -1 . "\n";
"#), vec!["256","256","1","0.5"]);
}

#[test]
fn log_functions_natural_10_2() {
    assert_eq!(run_prints(r#"<?php
echo round(log(M_E), 4) . "\n";
echo round(log10(1000), 4) . "\n";
echo round(log2(1024), 4) . "\n";
echo round(log(8, 2), 4) . "\n";
"#), vec!["1","3","10","3"]);
}

#[test]
fn sqrt_and_hypot() {
    assert_eq!(run_prints(r#"<?php
echo sqrt(144) . "\n";
echo sqrt(2) . "\n";
echo hypot(3, 4) . "\n";
echo hypot(5, 12) . "\n";
"#), vec!["12","1.4142135623731","5","13"]);
}

#[test]
fn pi_constant() {
    assert_eq!(run_prints(r#"<?php
echo round(pi(), 5) . "\n";
echo round(M_PI, 5) . "\n";
echo (pi() === M_PI) ? 'equal' : 'not equal';
echo "\n";
"#), vec!["3.14159","3.14159","equal"]);
}

#[test]
fn trigonometric_functions() {
    assert_eq!(run_prints(r#"<?php
echo round(sin(0), 4) . "\n";
echo round(cos(0), 4) . "\n";
echo round(tan(M_PI / 4), 4) . "\n";
echo round(sin(M_PI / 2), 4) . "\n";
echo round(cos(M_PI), 4) . "\n";
echo round(atan2(1, 1), 4) . "\n";
"#), vec!["0","1","1","1","-1","0.7854"]);
}

#[test]
fn inverse_trig_functions() {
    assert_eq!(run_prints(r#"<?php
echo round(asin(1.0), 4) . "\n";
echo round(acos(1.0), 4) . "\n";
echo round(atan(1.0), 4) . "\n";
echo round(rad2deg(asin(0.5)), 2) . "\n";
"#), vec!["1.5708","0","0.7854","30"]);
}

#[test]
fn hyperbolic_functions() {
    assert_eq!(run_prints(r#"<?php
echo round(sinh(0), 4) . "\n";
echo round(cosh(0), 4) . "\n";
echo round(tanh(0), 4) . "\n";
echo round(sinh(1), 4) . "\n";
echo round(cosh(1), 4) . "\n";
echo round(tanh(1), 4) . "\n";
"#), vec!["0","1","0","1.1752","1.5431","0.7616"]);
}

#[test]
fn numeric_string_comparison() {
    assert_eq!(run_prints(r#"<?php
var_dump('1' == 1);
var_dump('01' == '1');
var_dump('10' == '1e1');
var_dump('0' == false);
var_dump('0' === false);
"#), vec!["bool(true)","bool(true)","bool(true)","bool(true)","bool(false)"]);
}
