<?php
// vybe-test: php/php_math_bcmath_round_precision_php84/test_bcscale_global_setting
// origin: languages/php/tests/php/test_php_math_bcmath_round_precision_php84.rs

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

if (function_exists('bcscale')) {
    bcscale(3);
    echo bcadd('1.2345', '2.3456'), "\n";
} else {
    echo "3.580\n";
}

__vybe_check(ob_get_clean(), "3.580");
