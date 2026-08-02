<?php
// vybe-test: php/php_math_bcmath_number_precision/test_php_math_round_modes_half_up_even_down
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

$val = 2.5;
echo round($val, 0, PHP_ROUND_HALF_UP) . " | ";
echo round($val, 0, PHP_ROUND_HALF_EVEN) . " | ";
echo round($val, 0, PHP_ROUND_HALF_DOWN);

__vybe_check(ob_get_clean(), "3 | 2 | 2");
