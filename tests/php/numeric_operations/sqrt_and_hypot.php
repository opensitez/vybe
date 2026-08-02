<?php
// vybe-test: php/numeric_operations/sqrt_and_hypot
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

echo sqrt(144) . "\n";
echo sqrt(2) . "\n";
echo hypot(3, 4) . "\n";
echo hypot(5, 12) . "\n";

__vybe_check(ob_get_clean(), "12\n1.4142135623731\n5\n13");
