<?php
// vybe-test: php/numeric_operations/abs_on_negative_float
// origin: languages/php/tests/php/test_numeric_operations.rs

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

echo abs(-3.14) . "\n";
echo abs(3.14) . "\n";
echo abs(-0) . "\n";
echo abs(-PHP_INT_MAX) . "\n";

__vybe_check(ob_get_clean(), "3.14\n3.14\n0\n9223372036854775807");
