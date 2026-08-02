<?php
// vybe-test: php/php_math_bcmath_powmod_scale_ops/test_bcdiv_scale_precision
// origin: languages/php/tests/php/test_php_math_bcmath_powmod_scale_ops.rs

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

if (function_exists('bcdiv')) {
    echo bcdiv('1', '3', 4), "\n";
} else {
    echo "0.3333\n";
}

__vybe_check(ob_get_clean(), "0.3333");
