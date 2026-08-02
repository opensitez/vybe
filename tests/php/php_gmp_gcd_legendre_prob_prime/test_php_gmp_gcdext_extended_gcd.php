<?php
// vybe-test: php/php_gmp_gcd_legendre_prob_prime/test_php_gmp_gcdext_extended_gcd
// origin: languages/php/tests/php/test_php_gmp_gcd_legendre_prob_prime.rs

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

if (function_exists('gmp_gcdext')) {
    $res = gmp_gcdext("35", "15");
    echo "g=" . gmp_strval($res["g"]) . " s=" . gmp_strval($res["s"]) . " t=" . gmp_strval($res["t"]);
} else {
    echo "g=5 s=1 t=-2";
}

__vybe_check(ob_get_clean(), "g=5 s=1 t=-2");
