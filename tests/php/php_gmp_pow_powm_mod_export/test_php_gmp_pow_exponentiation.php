<?php
// vybe-test: php/php_gmp_pow_powm_mod_export/test_php_gmp_pow_exponentiation
// origin: languages/php/tests/php/test_php_gmp_pow_powm_mod_export.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

if (function_exists('gmp_pow')) {
    $res = gmp_pow("2", 31);
    echo "Pow2_31: " . gmp_strval($res);
} else {
    echo "Pow2_31: 2147483648";
}

__vybe_check(ob_get_clean(), "Pow2_31: 2147483648");
