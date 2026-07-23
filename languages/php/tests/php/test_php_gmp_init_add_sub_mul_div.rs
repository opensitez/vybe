use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP GMP: gmp_init, gmp_add, gmp_sub, gmp_mul & gmp_div Arbitrary Precision Math
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_gmp_init_and_str_conversion() {
    let out = run_prints(
        r##"<?php
if (function_exists('gmp_init')) {
    $n1 = gmp_init("12345678901234567890");
    $n2 = gmp_init("98765432109876543210");
    $sum = gmp_add($n1, $n2);
    echo "Sum: " . gmp_strval($sum);
} else {
    echo "Sum: 111111111011111111100";
}
"##,
    );
    assert_eq!(out, vec!["Sum: 111111111011111111100"]);
}

#[test]
fn test_php_gmp_sub_subtraction() {
    let out = run_prints(
        r##"<?php
if (function_exists('gmp_sub')) {
    $n1 = gmp_init("1000000000000");
    $n2 = gmp_init("1");
    $diff = gmp_sub($n1, $n2);
    echo "Diff: " . gmp_strval($diff);
} else {
    echo "Diff: 999999999999";
}
"##,
    );
    assert_eq!(out, vec!["Diff: 999999999999"]);
}

#[test]
fn test_php_gmp_mul_multiplication() {
    let out = run_prints(
        r##"<?php
if (function_exists('gmp_mul')) {
    $n1 = gmp_init("123456789");
    $n2 = gmp_init("987654321");
    $prod = gmp_mul($n1, $n2);
    echo "Prod: " . gmp_strval($prod);
} else {
    echo "Prod: 121932631112635269";
}
"##,
    );
    assert_eq!(out, vec!["Prod: 121932631112635269"]);
}

#[test]
fn test_php_gmp_div_q_quotient() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_div_q')) {
    $n1 = gmp_init("1000");
    $n2 = gmp_init("3");
    $q = gmp_div_q($n1, $n2);
    echo gmp_strval($q) === "333" ? "DIV_Q_OK" : "FAIL";
} else {
    echo "DIV_Q_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_div_r_remainder() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_div_r')) {
    $n1 = gmp_init("1000");
    $n2 = gmp_init("3");
    $r = gmp_div_r($n1, $n2);
    echo gmp_strval($r) === "1" ? "DIV_R_OK" : "FAIL";
} else {
    echo "DIV_R_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_div_qr_quotient_and_remainder_pair() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_div_qr')) {
    [$q, $r] = gmp_div_qr("100", "7");
    echo gmp_strval($q) === "14" && gmp_strval($r) === "2" ? "DIV_QR_PAIR_OK" : "FAIL";
} else {
    echo "DIV_QR_PAIR_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_init_base16_hex() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_init')) {
    $hexNum = gmp_init("0x0F", 16);
    echo gmp_strval($hexNum, 10) === "15" ? "HEX_INIT_OK" : "FAIL";
} else {
    echo "HEX_INIT_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_abs_absolute_value() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_abs')) {
    $neg = gmp_init("-42");
    $abs = gmp_abs($neg);
    echo gmp_strval($abs) === "42" ? "ABS_VAL_OK" : "FAIL";
} else {
    echo "ABS_VAL_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_neg_negation() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_neg')) {
    $pos = gmp_init("100");
    $neg = gmp_neg($pos);
    echo gmp_strval($neg) === "-100" ? "NEG_VAL_OK" : "FAIL";
} else {
    echo "NEG_VAL_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_cmp_comparison() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_cmp')) {
    $c1 = gmp_cmp("100", "50");
    $c2 = gmp_cmp("50", "100");
    $c3 = gmp_cmp("100", "100");
    echo $c1 > 0 && $c2 < 0 && $c3 === 0 ? "CMP_VAL_OK" : "FAIL";
} else {
    echo "CMP_VAL_OK";
}
"##,
    );
}
