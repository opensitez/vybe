<?php
// vybe-test: php/numeric_operations/integer_overflow_to_float
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

$max = PHP_INT_MAX;
$overflow = $max + 1;
echo is_float($overflow) ? 'float' : 'int';
echo "\n";
echo is_int($max) ? 'int' : 'float';
echo "\n";

__vybe_check(ob_get_clean(), "float\nint");
