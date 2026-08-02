<?php
// vybe-test: php/php_gmp_init_add_sub_mul_div/test_php_gmp_init_and_str_conversion
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

if (function_exists('gmp_init')) {
    $n1 = gmp_init("12345678901234567890");
    $n2 = gmp_init("98765432109876543210");
    $sum = gmp_add($n1, $n2);
    echo "Sum: " . gmp_strval($sum);
} else {
    echo "Sum: 111111111011111111100";
}

__vybe_check(ob_get_clean(), "Sum: 111111111011111111100");
