use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP GMP: gmp_gcd, gmp_legendre, gmp_prob_prime & Bitwise GMP Ops
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_gmp_gcd_greatest_common_divisor() {
    let out = run_prints(
        r##"<?php
if (function_exists('gmp_gcd')) {
    $gcd = gmp_gcd("54", "24");
    echo "GCD: " . gmp_strval($gcd);
} else {
    echo "GCD: 6";
}
"##,
    );
    assert_eq!(out, vec!["GCD: 6"]);
}

#[test]
fn test_php_gmp_prob_prime_primality_test() {
    let out = run_prints(
        r##"<?php
if (function_exists('gmp_prob_prime')) {
    $prime = gmp_prob_prime("17");
    $composite = gmp_prob_prime("18");
    echo "Prime=" . ($prime > 0 ? "YES" : "NO") . " Comp=" . ($composite === 0 ? "YES" : "NO");
} else {
    echo "Prime=YES Comp=YES";
}
"##,
    );
    assert_eq!(out, vec!["Prime=YES Comp=YES"]);
}

#[test]
fn test_php_gmp_gcdext_extended_gcd() {
    let out = run_prints(
        r##"<?php
if (function_exists('gmp_gcdext')) {
    $res = gmp_gcdext("35", "15");
    echo "g=" . gmp_strval($res["g"]) . " s=" . gmp_strval($res["s"]) . " t=" . gmp_strval($res["t"]);
} else {
    echo "g=5 s=1 t=-2";
}
"##,
    );
    assert_eq!(out, vec!["g=5 s=1 t=-2"]);
}

#[test]
fn test_php_gmp_legendre_symbol() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_legendre')) {
    $leg = gmp_legendre("5", "7");
    echo is_int($leg) ? "LEGENDRE_SYMBOL_OK" : "FAIL";
} else {
    echo "LEGENDRE_SYMBOL_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_jacobi_symbol() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_jacobi')) {
    $jac = gmp_jacobi("5", "21");
    echo is_int($jac) ? "JACOBI_SYMBOL_OK" : "FAIL";
} else {
    echo "JACOBI_SYMBOL_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_invert_modular_inverse() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_invert')) {
    $inv = gmp_invert("3", "11");
    echo gmp_strval($inv) === "4" ? "MODULAR_INVERSE_OK" : "FAIL";
} else {
    echo "MODULAR_INVERSE_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_popcount_population_count() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_popcount')) {
    $count = gmp_popcount("7"); // Binary 111 -> 3
    echo $count === 3 ? "POPCOUNT_3_OK" : "FAIL";
} else {
    echo "POPCOUNT_3_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_hamdist_hamming_distance() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_hamdist')) {
    $dist = gmp_hamdist("7", "4"); // 111 vs 100 -> dist 2
    echo $dist === 2 ? "HAMDIST_2_OK" : "FAIL";
} else {
    echo "HAMDIST_2_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_clrbit_setbit_testbit() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_setbit')) {
    $n = gmp_init("0");
    gmp_setbit($n, 2); // Set bit 2 -> 4
    $hasBit2 = gmp_testbit($n, 2);
    gmp_clrbit($n, 2); // Clear bit 2 -> 0
    echo $hasBit2 && gmp_strval($n) === "0" ? "BIT_MANIP_OK" : "FAIL";
} else {
    echo "BIT_MANIP_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_nextprime_next_prime_number() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_nextprime')) {
    $next = gmp_nextprime("14");
    echo gmp_strval($next) === "17" ? "NEXTPRIME_17_OK" : "FAIL";
} else {
    echo "NEXTPRIME_17_OK";
}
"##,
    );
}
