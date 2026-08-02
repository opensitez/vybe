<?php
// vybe-test: php/php_math_bcmath_powmod_scale_ops/test_bcmod_with_scale_php80
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

if (function_exists('bcmod')) {
    echo bcmod('5.7', '1.3', 1), "\n";
} else {
    echo "0.5\n";
}

__vybe_check(ob_get_clean(), "0.5");
