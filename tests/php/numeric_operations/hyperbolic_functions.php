<?php
// vybe-test: php/numeric_operations/hyperbolic_functions
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

echo round(sinh(0), 4) . "\n";
echo round(cosh(0), 4) . "\n";
echo round(tanh(0), 4) . "\n";
echo round(sinh(1), 4) . "\n";
echo round(cosh(1), 4) . "\n";
echo round(tanh(1), 4) . "\n";

__vybe_check(ob_get_clean(), "0\n1\n0\n1.1752\n1.5431\n0.7616");
