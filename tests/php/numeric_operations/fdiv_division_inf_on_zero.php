<?php
// vybe-test: php/numeric_operations/fdiv_division_inf_on_zero
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

echo fdiv(10, 0) === INF ? 'INF' : 'other';
echo "\n";
echo fdiv(-10, 0) === -INF ? '-INF' : 'other';
echo "\n";
echo fdiv(0, 0) !== fdiv(0, 0) ? 'NAN' : 'other';
echo "\n";
echo fdiv(10, 2) . "\n";

__vybe_check(ob_get_clean(), "INF\n-INF\nNAN\n5");
