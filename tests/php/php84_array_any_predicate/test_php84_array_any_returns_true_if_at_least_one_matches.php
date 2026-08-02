<?php
// vybe-test: php/php84_array_any_predicate/test_php84_array_any_returns_true_if_at_least_one_matches
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

$nums = [1, 3, 5, 8, 9];
if (function_exists('array_any')) {
    $hasEven = array_any($nums, fn($n) => $n % 2 === 0);
    echo $hasEven ? "HAS_EVEN_TRUE" : "FALSE";
} else {
    $hasEven = false;
    foreach ($nums as $n) { if ($n % 2 === 0) { $hasEven = true; break; } }
    echo $hasEven ? "HAS_EVEN_TRUE" : "FALSE";
}

__vybe_check(ob_get_clean(), "HAS_EVEN_TRUE");
