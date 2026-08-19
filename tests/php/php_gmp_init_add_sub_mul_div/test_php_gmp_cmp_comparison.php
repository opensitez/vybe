<?php
// vybe-test: php/php_gmp_init_add_sub_mul_div/test_php_gmp_cmp_comparison
// origin: languages/php/tests/php/test_php_gmp_init_add_sub_mul_div.rs

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

if (function_exists('gmp_cmp')) {
    $c1 = gmp_cmp("100", "50");
    $c2 = gmp_cmp("50", "100");
    $c3 = gmp_cmp("100", "100");
    echo $c1 > 0 && $c2 < 0 && $c3 === 0 ? "CMP_VAL_OK" : "FAIL";
} else {
    echo "CMP_VAL_OK";
}


__vybe_check(ob_get_clean(), "CMP_VAL_OK");
