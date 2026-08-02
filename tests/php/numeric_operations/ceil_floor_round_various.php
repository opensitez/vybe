<?php
// vybe-test: php/numeric_operations/ceil_floor_round_various
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

echo ceil(4.1) . "\n";
echo ceil(-4.1) . "\n";
echo floor(4.9) . "\n";
echo floor(-4.9) . "\n";
echo round(4.5) . "\n";
echo round(4.55, 1) . "\n";
echo round(-4.5) . "\n";

__vybe_check(ob_get_clean(), "5\n-4\n4\n-5\n5\n4.6\n-5");
