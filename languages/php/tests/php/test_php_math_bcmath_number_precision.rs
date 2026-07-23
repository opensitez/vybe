use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Math & BCMath High Precision — bcadd, bcsub, bcmul, bcdiv, bcmod, bcpow, abs, round, ceil, floor, min, max, pow, fmod, intdiv
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_bcmath_arbitrary_precision_addition() {
    let out = run_prints(
        r#"<?php
$a = "1.234567890123456789";
$b = "9.876543210987654321";
echo bcadd($a, $b, 10);
"#,
    );
    assert_eq!(out, vec!["11.1111111011"]);
}

#[test]
fn test_php_bcmath_subtraction_and_multiplication() {
    let out = run_prints(
        r#"<?php
$sub = bcsub("10.50", "3.25", 2);
$mul = bcmul("2.5", "4.0", 2);
echo "$sub | $mul";
"#,
    );
    assert_eq!(out, vec!["7.25 | 10.00"]);
}

#[test]
fn test_php_math_round_modes_half_up_even_down() {
    let out = run_prints(
        r#"<?php
$val = 2.5;
echo round($val, 0, PHP_ROUND_HALF_UP) . " | ";
echo round($val, 0, PHP_ROUND_HALF_EVEN) . " | ";
echo round($val, 0, PHP_ROUND_HALF_DOWN);
"#,
    );
    assert_eq!(out, vec!["3 | 2 | 2"]);
}

#[test]
fn test_php_intdiv_and_fmod_arithmetic() {
    let out = run_prints(
        r#"<?php
$quotient = intdiv(10, 3);
$remainder = fmod(10.5, 3.0);
echo "intdiv=$quotient fmod=$remainder";
"#,
    );
    assert_eq!(out, vec!["intdiv=3 fmod=1.5"]);
}

#[test]
fn test_php_min_max_array_and_varargs() {
    let out = run_prints(
        r#"<?php
$minVal = min([10, 5, 20, 3]);
$maxVal = max(10, 50, 20, 30);
echo "Min=$minVal Max=$maxVal";
"#,
    );
    assert_eq!(out, vec!["Min=3 Max=50"]);
}

#[test]
fn test_php_bcsqrt_and_bcpow() {
    compile_ok(
        r#"<?php
$sqrt = bcsqrt("2", 6);
$pow = bcpow("2", "10", 0);
echo "sqrt2=$sqrt pow=$pow";
"#,
    );
}

#[test]
fn test_php_bcscale_default_setting() {
    compile_ok(
        r#"<?php
bcscale(4);
echo bcadd("1.11111", "2.22222");
"#,
    );
}

#[test]
fn test_php_hypot_deg2rad_rad2deg() {
    compile_ok(
        r#"<?php
$h = hypot(3, 4);
$rad = deg2rad(180);
$deg = rad2deg(M_PI);
echo "hypot=$h rad=" . round($rad, 2) . " deg=" . round($deg, 0);
"#,
    );
}

#[test]
fn test_php_is_nan_is_infinite_checks() {
    compile_ok(
        r#"<?php
$nan = acos(8);
$inf = log(0);
echo is_nan($nan) ? "NAN" : "NOT_NAN";
echo is_infinite($inf) ? " INF" : " NOT_INF";
"#,
    );
}

#[test]
fn test_php_base_convert_hex_dec_bin() {
    compile_ok(
        r#"<?php
$hex = "FF";
$dec = base_convert($hex, 16, 10);
$bin = base_convert($dec, 10, 2);
echo "Dec=$dec Bin=$bin";
"#,
    );
}
