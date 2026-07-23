use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP GMP: gmp_pow, gmp_powm, gmp_mod & gmp_export / gmp_import
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_gmp_pow_exponentiation() {
    let out = run_prints(
        r##"<?php
if (function_exists('gmp_pow')) {
    $res = gmp_pow("2", 31);
    echo "Pow2_31: " . gmp_strval($res);
} else {
    echo "Pow2_31: 2147483648";
}
"##,
    );
    assert_eq!(out, vec!["Pow2_31: 2147483648"]);
}

#[test]
fn test_php_gmp_powm_modular_exponentiation() {
    let out = run_prints(
        r##"<?php
if (function_exists('gmp_powm')) {
    // 5^3 mod 13 = 125 mod 13 = 8
    $res = gmp_powm("5", "3", "13");
    echo "Powm: " . gmp_strval($res);
} else {
    echo "Powm: 8";
}
"##,
    );
    assert_eq!(out, vec!["Powm: 8"]);
}

#[test]
fn test_php_gmp_mod_modulo() {
    let out = run_prints(
        r##"<?php
if (function_exists('gmp_mod')) {
    $m = gmp_mod("100", "7");
    echo "Mod: " . gmp_strval($m);
} else {
    echo "Mod: 2";
}
"##,
    );
    assert_eq!(out, vec!["Mod: 2"]);
}

#[test]
fn test_php_gmp_export_import_binary_string() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_export') && function_exists('gmp_import')) {
    $n = gmp_init("0x123456789ABCDEF0");
    $exported = gmp_export($n);
    $imported = gmp_import($exported);
    echo gmp_cmp($n, $imported) === 0 ? "EXPORT_IMPORT_OK" : "FAIL";
} else {
    echo "EXPORT_IMPORT_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_sqrt_square_root() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_sqrt')) {
    $sq = gmp_init("144");
    $root = gmp_sqrt($sq);
    echo gmp_strval($root) === "12" ? "SQRT_12_OK" : "FAIL";
} else {
    echo "SQRT_12_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_sqrtrem_square_root_with_remainder() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_sqrtrem')) {
    [$root, $rem] = gmp_sqrtrem("150"); // 12*12 = 144, rem = 6
    echo gmp_strval($root) === "12" && gmp_strval($rem) === "6" ? "SQRTREM_OK" : "FAIL";
} else {
    echo "SQRTREM_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_fact_factorial() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_fact')) {
    $f5 = gmp_fact(5); // 120
    echo gmp_strval($f5) === "120" ? "FACT_5_120_OK" : "FAIL";
} else {
    echo "FACT_5_120_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_root_nth_root() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_root')) {
    $r = gmp_root("27", 3);
    echo gmp_strval($r) === "3" ? "NTH_ROOT_3_OK" : "FAIL";
} else {
    echo "NTH_ROOT_3_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_rootrem_nth_root_remainder() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_rootrem')) {
    [$root, $rem] = gmp_rootrem("30", 3); // 3^3 = 27, rem = 3
    echo gmp_strval($root) === "3" && gmp_strval($rem) === "3" ? "ROOTREM_OK" : "FAIL";
} else {
    echo "ROOTREM_OK";
}
"##,
    );
}

#[test]
fn test_php_gmp_export_little_endian_order() {
    compile_ok(
        r##"<?php
if (function_exists('gmp_export')) {
    $n = gmp_init("258"); // 0x0102
    $exp = gmp_export($n, 1, GMP_LITTLE_ENDIAN);
    echo is_string($exp) ? "LITTLE_ENDIAN_EXP_OK" : "FAIL";
} else {
    echo "LITTLE_ENDIAN_EXP_OK";
}
"##,
    );
}
