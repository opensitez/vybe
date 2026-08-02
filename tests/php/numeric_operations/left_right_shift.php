<?php
// vybe-test: php/numeric_operations/left_right_shift
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

echo (1 << 4) . "\n";
echo (256 >> 3) . "\n";
echo (0b0001 << 3) . "\n";
echo (0b10000 >> 2) . "\n";

__vybe_check(ob_get_clean(), "16\n32\n8\n4");
