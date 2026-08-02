<?php
// vybe-test: php/numeric_operations/base_convert_between_bases
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

echo base_convert('ff', 16, 10) . "\n";
echo base_convert('255', 10, 16) . "\n";
echo base_convert('11111111', 2, 10) . "\n";
echo base_convert('377', 8, 10) . "\n";

__vybe_check(ob_get_clean(), "255\nff\n255\n255");
