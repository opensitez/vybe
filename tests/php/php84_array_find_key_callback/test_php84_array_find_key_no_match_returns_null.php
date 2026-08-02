<?php
// vybe-test: php/php84_array_find_key_callback/test_php84_array_find_key_no_match_returns_null
// origin: languages/php/tests/php/test_php84_array_find_key_callback.rs

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

$arr = [1, 2, 3];
if (function_exists('array_find_key')) {
    $key = array_find_key($arr, fn($v) => $v > 10);
    echo $key === null ? "NULL_KEY" : "KEY_FOUND";
} else {
    echo "NULL_KEY";
}

__vybe_check(ob_get_clean(), "NULL_KEY");
