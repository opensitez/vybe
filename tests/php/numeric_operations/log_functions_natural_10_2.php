<?php
// vybe-test: php/numeric_operations/log_functions_natural_10_2
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

echo round(log(M_E), 4) . "\n";
echo round(log10(1000), 4) . "\n";
echo round(log2(1024), 4) . "\n";
echo round(log(8, 2), 4) . "\n";

__vybe_check(ob_get_clean(), "1\n3\n10\n3");
