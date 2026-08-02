<?php
// vybe-test: php/php_math_bcmath_number_precision/test_php_bcmath_arbitrary_precision_addition
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

$a = "1.234567890123456789";
$b = "9.876543210987654321";
echo bcadd($a, $b, 10);

__vybe_check(ob_get_clean(), "11.1111111011");
