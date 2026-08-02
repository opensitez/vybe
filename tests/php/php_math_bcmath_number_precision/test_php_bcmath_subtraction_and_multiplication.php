<?php
// vybe-test: php/php_math_bcmath_number_precision/test_php_bcmath_subtraction_and_multiplication
// origin: languages/php/tests/php/test_php_math_bcmath_number_precision.rs

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

$sub = bcsub("10.50", "3.25", 2);
$mul = bcmul("2.5", "4.0", 2);
echo "$sub | $mul";

__vybe_check(ob_get_clean(), "7.25 | 10.00");
