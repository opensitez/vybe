<?php
// vybe-test: php/numeric_operations/inverse_trig_functions
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

echo round(asin(1.0), 4) . "\n";
echo round(acos(1.0), 4) . "\n";
echo round(atan(1.0), 4) . "\n";
echo round(rad2deg(asin(0.5)), 2) . "\n";

__vybe_check(ob_get_clean(), "1.5708\n0\n0.7854\n30");
