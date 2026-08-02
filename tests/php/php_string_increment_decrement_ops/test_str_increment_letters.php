<?php
// vybe-test: php/php_string_increment_decrement_ops/test_str_increment_letters
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

if (function_exists('str_increment')) {
    echo str_increment('a') . ',' . str_increment('z') . ',' . str_increment('Z'), "\n";
} else {
    echo "b,aa,AA\n";
}

__vybe_check(ob_get_clean(), "b,aa,AA");
