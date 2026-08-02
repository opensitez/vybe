<?php
// vybe-test: php/php_math_intdiv_fmod_edge_cases/test_intdiv_division_by_zero_throws
// origin: languages/php/tests/php/test_php_math_intdiv_fmod_edge_cases.rs

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

try {
    intdiv(5, 0);
    echo "no_error\n";
} catch (DivisionByZeroError $e) {
    echo "div_zero_error\n";
}

__vybe_check(ob_get_clean(), "div_zero_error");
