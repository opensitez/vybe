use super::helpers::run_prints;

// ── Integer operations ────────────────────────────────────────

#[test] fn integer_division_truncates() {
    assert_eq!(run_prints(r#"<?php echo intdiv(7, 2); "#), vec!["3"]);
}
#[test] fn modulo_sign_follows_dividend() {
    assert_eq!(run_prints(r#"<?php echo -7 % 3; "#), vec!["-1"]);
}
#[test] fn exponentiation_operator() {
    assert_eq!(run_prints(r#"<?php echo 2 ** 10; "#), vec!["1024"]);
}
#[test] fn large_number_precision() {
    assert_eq!(run_prints(r#"<?php echo PHP_INT_MAX + 1 > PHP_INT_MAX ? 'overflow' : 'same'; "#), vec!["overflow"]);
}

// ── Float operations ──────────────────────────────────────────

#[test] fn float_precision_comparison() {
    assert_eq!(run_prints(r#"<?php echo abs(0.1 + 0.2 - 0.3) < PHP_FLOAT_EPSILON ? 'equal' : 'diff'; "#), vec!["equal"]);
}
#[test] fn float_nan_properties() {
    assert_eq!(run_prints(r#"<?php $n = NAN; echo is_nan($n) ? 'nan' : 'not'; echo ($n === $n) ? 'eq' : 'neq'; "#), vec!["nanneq"]);
}
#[test] fn float_inf_operations() {
    assert_eq!(run_prints(r#"<?php echo INF + 1 === INF ? 'same' : 'diff'; echo -INF < 0 ? 'neg' : 'pos'; "#), vec!["sameneg"]);
}
#[test] fn scientific_notation_float() {
    assert_eq!(run_prints(r#"<?php echo 1.5e3; "#), vec!["1500"]);
}

// ── Number functions ──────────────────────────────────────────

#[test] fn abs_of_various() {
    assert_eq!(run_prints(r#"<?php echo abs(-5) . ',' . abs(5) . ',' . abs(-3.14); "#), vec!["5,5,3.14"]);
}
#[test] fn max_min_with_various_types() {
    assert_eq!(run_prints(r#"<?php echo max('10', 9, 11) . ',' . min(0, -1, 1); "#), vec!["11,-1"]);
}
#[test] fn round_various_precisions() {
    assert_eq!(run_prints(r#"<?php echo round(1.555, 2) . ',' . round(1234.5, -2); "#), vec!["1.56,1200"]);
}

// ── Random number generation ──────────────────────────────────

#[test] fn rand_range() {
    assert_eq!(run_prints(r#"<?php
$n = rand(1, 10);
echo ($n >= 1 && $n <= 10) ? 'in_range' : 'out';
"#), vec!["in_range"]);
}
#[test] fn mt_rand_consistent_range() {
    assert_eq!(run_prints(r#"<?php
$n = mt_rand(100, 200);
echo ($n >= 100 && $n <= 200) ? 'ok' : 'fail';
"#), vec!["ok"]);
}
#[test] fn random_int_range() {
    assert_eq!(run_prints(r#"<?php
$n = random_int(1, 6);
echo ($n >= 1 && $n <= 6) ? 'ok' : 'fail';
"#), vec!["ok"]);
}

// ── Number formatting ─────────────────────────────────────────

#[test] fn number_format_zero_decimals() {
    assert_eq!(run_prints(r#"<?php echo number_format(1234567); "#), vec!["1,234,567"]);
}
#[test] fn decoct_decbin_dechex() {
    assert_eq!(run_prints(r#"<?php echo decoct(8) . ',' . decbin(10) . ',' . dechex(255); "#), vec!["10,1010,ff"]);
}
#[test] fn octdec_bindec_hexdec() {
    assert_eq!(run_prints(r#"<?php echo octdec('10') . ',' . bindec('1010') . ',' . hexdec('ff'); "#), vec!["8,10,255"]);
}

// ── GMP (arbitrary precision) ────────────────────────────────

#[test] fn bcmath_add() {
    assert_eq!(run_prints(r#"<?php echo bcadd('999999999999999999', '1'); "#), vec!["1000000000000000000"]);
}
#[test] fn bcmath_mul() {
    assert_eq!(run_prints(r#"<?php echo bcmul('123456789', '987654321'); "#), vec!["121932631112635269"]);
}
#[test] fn bcmath_div() {
    assert_eq!(run_prints(r#"<?php echo bcdiv('10', '3', 4); "#), vec!["3.3333"]);
}
#[test] fn bcmath_pow() {
    assert_eq!(run_prints(r#"<?php echo bcpow('2', '64'); "#), vec!["18446744073709551616"]);
}
#[test] fn bcmath_comp() {
    assert_eq!(run_prints(r#"<?php echo bccomp('1.23456789', '1.23456790'); "#), vec!["-1"]);
}
#[test] fn bcmath_scale() {
    assert_eq!(run_prints(r#"<?php bcscale(2); echo bcadd('1', '2'); "#), vec!["3.00"]);
}
