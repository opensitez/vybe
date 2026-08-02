<?php
// vybe-test: php/fdiv_arithmetic/integer_overflow_becomes_float
// origin: languages/php/tests/php/test_fdiv_arithmetic.rs

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

$r = PHP_INT_MAX + 1;
echo is_float($r) ? 'float' : 'int';

__vybe_check(ob_get_clean(), "float");
