<?php
// vybe-test: php/php84_array_find_callback/test_php84_array_find_no_match_returns_null
// origin: languages/php/tests/php/test_php84_array_find_callback.rs

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

$arr = ["apple", "banana", "cherry"];
if (function_exists('array_find')) {
    $res = array_find($arr, fn($item) => str_starts_with($item, "z"));
    echo $res === null ? "NULL_MATCH" : "FOUND";
} else {
    echo "NULL_MATCH";
}

__vybe_check(ob_get_clean(), "NULL_MATCH");
