use super::helpers::compile_ok;

// ── bcadd ────────────────────────────────────────────────────

#[test]
fn bcadd_basic() {
    compile_ok(
        r#"<?php
echo bcadd('10', '5');
echo bcadd('10', '5', 2);
"#,
    );
}

#[test]
fn bcadd_large_numbers() {
    compile_ok(
        r#"<?php
$a = '99999999999999999999';
$b = '1';
echo bcadd($a, $b);
"#,
    );
}

#[test]
fn bcadd_decimals() {
    compile_ok(
        r#"<?php
echo bcadd('1.5', '2.3', 1);
echo bcadd('0.1', '0.2', 1);
"#,
    );
}

#[test]
fn bcadd_negative() {
    compile_ok(
        r#"<?php
echo bcadd('-5', '3');
echo bcadd('-5', '-3');
"#,
    );
}

// ── bcsub ────────────────────────────────────────────────────

#[test]
fn bcsub_basic() {
    compile_ok(
        r#"<?php
echo bcsub('10', '3');
echo bcsub('10', '3', 2);
"#,
    );
}

#[test]
fn bcsub_large() {
    compile_ok(
        r#"<?php
echo bcsub('100000000000000000000', '1');
"#,
    );
}

#[test]
fn bcsub_negative_result() {
    compile_ok(
        r#"<?php
echo bcsub('3', '10');
echo bcsub('3.5', '10.2', 1);
"#,
    );
}

// ── bcmul ────────────────────────────────────────────────────

#[test]
fn bcmul_basic() {
    compile_ok(
        r#"<?php
echo bcmul('6', '7');
echo bcmul('3.14', '2', 2);
"#,
    );
}

#[test]
fn bcmul_large() {
    compile_ok(
        r#"<?php
echo bcmul('999999999999', '999999999999');
"#,
    );
}

#[test]
fn bcmul_decimal_precision() {
    compile_ok(
        r#"<?php
echo bcmul('1.23456', '9.87654', 5);
"#,
    );
}

#[test]
fn bcmul_by_zero() {
    compile_ok(
        r#"<?php
echo bcmul('12345678901234567890', '0');
"#,
    );
}

// ── bcdiv ────────────────────────────────────────────────────

#[test]
fn bcdiv_basic() {
    compile_ok(
        r#"<?php
echo bcdiv('10', '3', 5);
echo bcdiv('1', '3', 10);
"#,
    );
}

#[test]
fn bcdiv_exact() {
    compile_ok(
        r#"<?php
echo bcdiv('100', '4', 2);
echo bcdiv('22', '7', 6);
"#,
    );
}

#[test]
fn bcdiv_large_dividend() {
    compile_ok(
        r#"<?php
echo bcdiv('123456789012345678', '9', 0);
"#,
    );
}

// ── bcmod ────────────────────────────────────────────────────

#[test]
fn bcmod_basic() {
    compile_ok(
        r#"<?php
echo bcmod('10', '3');
echo bcmod('17', '5');
"#,
    );
}

#[test]
fn bcmod_large() {
    compile_ok(
        r#"<?php
echo bcmod('100000000000000000001', '7');
"#,
    );
}

#[test]
fn bcmod_decimal() {
    compile_ok(
        r#"<?php
echo bcmod('10.5', '3.2', 1);
"#,
    );
}

// ── bcpow ────────────────────────────────────────────────────

#[test]
fn bcpow_basic() {
    compile_ok(
        r#"<?php
echo bcpow('2', '10');
echo bcpow('3', '5');
"#,
    );
}

#[test]
fn bcpow_large_exponent() {
    compile_ok(
        r#"<?php
echo bcpow('2', '64');
"#,
    );
}

#[test]
fn bcpow_decimal_base() {
    compile_ok(
        r#"<?php
echo bcpow('1.05', '12', 4);
"#,
    );
}

#[test]
fn bcpow_zero_exponent() {
    compile_ok(
        r#"<?php
echo bcpow('999', '0');
"#,
    );
}

// ── bcsqrt ───────────────────────────────────────────────────

#[test]
fn bcsqrt_basic() {
    compile_ok(
        r#"<?php
echo bcsqrt('2', 10);
echo bcsqrt('9', 2);
"#,
    );
}

#[test]
fn bcsqrt_perfect_square() {
    compile_ok(
        r#"<?php
echo bcsqrt('144', 0);
echo bcsqrt('10000', 0);
"#,
    );
}

#[test]
fn bcsqrt_high_precision() {
    compile_ok(
        r#"<?php
echo bcsqrt('2', 20);
"#,
    );
}

// ── bcscale ──────────────────────────────────────────────────

#[test]
fn bcscale_global() {
    compile_ok(
        r#"<?php
bcscale(4);
echo bcadd('1', '2');
echo bcdiv('1', '3');
bcscale(0); // reset
"#,
    );
}

#[test]
fn bcscale_returns_old() {
    compile_ok(
        r#"<?php
bcscale(0);
$old = bcscale(5);
echo is_bool($old) || is_int($old) ? 'ok' : 'fail';
bcscale(0);
"#,
    );
}

// ── bccomp ───────────────────────────────────────────────────

#[test]
fn bccomp_equal() {
    compile_ok(
        r#"<?php
echo bccomp('1.0', '1.00', 2);
echo bccomp('999', '999');
"#,
    );
}

#[test]
fn bccomp_less_greater() {
    compile_ok(
        r#"<?php
echo bccomp('1', '2');    // -1
echo bccomp('2', '1');    // 1
echo bccomp('1.5', '1.6', 1); // -1
"#,
    );
}

#[test]
fn bccomp_large_numbers() {
    compile_ok(
        r#"<?php
$big = '99999999999999999999';
$bigger = '100000000000000000000';
echo bccomp($big, $bigger);
echo bccomp($bigger, $big);
"#,
    );
}

// ── Practical bcmath patterns ─────────────────────────────────

#[test]
fn bcmath_financial_calculation() {
    compile_ok(
        r#"<?php
// Interest calculation without float precision loss
$principal = '1000.00';
$rate      = '0.05';
$periods   = '12';
$interest  = bcmul(bcmul($principal, $rate, 4), bcdiv($periods, '12', 4), 2);
$total     = bcadd($principal, $interest, 2);
echo $total;
"#,
    );
}

#[test]
fn bcmath_factorial_large() {
    compile_ok(
        r#"<?php
function bcfactorial(int $n): string {
    $result = '1';
    for ($i = 2; $i <= $n; $i++) {
        $result = bcmul($result, (string)$i);
    }
    return $result;
}
$f20 = bcfactorial(20);
echo strlen($f20) > 10 ? 'large factorial' : 'too small';
echo bccomp($f20, '2432902008176640000') >= 0 ? ':correct magnitude' : ':wrong';
"#,
    );
}

#[test]
fn bcmath_pi_digits() {
    compile_ok(
        r#"<?php
// Leibniz formula approximation using bc (illustrative)
$pi = '0';
$sign = '1';
for ($k = 0; $k < 100; $k++) {
    $term = bcdiv($sign, bcadd(bcmul('2', (string)$k), '1'), 20);
    $pi = bcadd($pi, $term, 20);
    $sign = bcmul($sign, '-1');
}
$pi = bcmul($pi, '4', 10);
echo bccomp($pi, '3.1') > 0 && bccomp($pi, '3.2') < 0 ? 'pi in range' : 'out of range';
"#,
    );
}

#[test]
fn bcmath_currency_rounding() {
    compile_ok(
        r#"<?php
// Avoid float rounding errors in currency
$prices = ['10.99', '5.49', '3.25', '0.01'];
$total = '0.00';
foreach ($prices as $p) {
    $total = bcadd($total, $p, 2);
}
echo $total;
"#,
    );
}

#[test]
fn bcmath_power_of_two_sequence() {
    compile_ok(
        r#"<?php
$result = [];
for ($i = 0; $i <= 10; $i++) {
    $result[] = bcpow('2', (string)$i);
}
echo implode(',', $result);
"#,
    );
}
