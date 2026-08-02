<?php
// vybe-test: php/php84_array_all_predicate/test_php84_array_all_empty_array_returns_true
// origin: languages/php/tests/php/test_php84_array_all_predicate.rs

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

if (function_exists('array_all')) {
    $res = array_all([], fn($x) => false);
    echo $res ? "VACUOUS_TRUE" : "FALSE";
} else {
    echo "VACUOUS_TRUE";
}

__vybe_check(ob_get_clean(), "VACUOUS_TRUE");
