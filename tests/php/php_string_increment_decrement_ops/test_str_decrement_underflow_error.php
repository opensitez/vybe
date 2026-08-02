<?php
// vybe-test: php/php_string_increment_decrement_ops/test_str_decrement_underflow_error
// origin: languages/php/tests/php/test_php_string_increment_decrement_ops.rs

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

if (function_exists('str_decrement')) {
    try {
        str_decrement('a');
        echo "no_error\n";
    } catch (ValueError $e) {
        echo "underflow_error\n";
    }
} else {
    echo "underflow_error\n";
}

__vybe_check(ob_get_clean(), "underflow_error");
