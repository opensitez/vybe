<?php
// vybe-test: php/php_math_bcmath_number_precision/test_php_min_max_array_and_varargs
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

$minVal = min([10, 5, 20, 3]);
$maxVal = max(10, 50, 20, 30);
echo "Min=$minVal Max=$maxVal";

__vybe_check(ob_get_clean(), "Min=3 Max=50");
