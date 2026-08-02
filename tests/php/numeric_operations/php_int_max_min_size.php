<?php
// vybe-test: php/numeric_operations/php_int_max_min_size
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

echo PHP_INT_SIZE . "\n";
echo (PHP_INT_MAX > 0) ? 'max_positive' : 'max_negative';
echo "\n";
echo (PHP_INT_MIN < 0) ? 'min_negative' : 'min_positive';
echo "\n";

__vybe_check(ob_get_clean(), "8\nmax_positive\nmin_negative");
