<?php
// vybe-test: php/numeric_operations/underscore_numeric_literal
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

$million = 1_000_000;
echo $million . "\n";
$pi = 3.141_592_653;
echo round($pi, 6) . "\n";
$hex = 0xFF_FF;
echo $hex . "\n";

__vybe_check(ob_get_clean(), "1000000\n3.141593\n65535");
