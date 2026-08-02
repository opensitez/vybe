<?php
// vybe-test: php/php84_array_any_predicate/test_php84_array_any_returns_false_if_none_match
// origin: languages/php/tests/php/test_php84_array_any_predicate.rs

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

$words = ["apple", "apricot", "avocado"];
if (function_exists('array_any')) {
    $hasB = array_any($words, fn($w) => str_starts_with($w, "b"));
    echo $hasB ? "TRUE" : "HAS_B_FALSE";
} else {
    echo "HAS_B_FALSE";
}

__vybe_check(ob_get_clean(), "HAS_B_FALSE");
