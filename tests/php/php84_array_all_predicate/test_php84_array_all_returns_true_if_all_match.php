<?php
// vybe-test: php/php84_array_all_predicate/test_php84_array_all_returns_true_if_all_match
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

$nums = [2, 4, 6, 8];
if (function_exists('array_all')) {
    $allEven = array_all($nums, fn($n) => $n % 2 === 0);
    echo $allEven ? "ALL_EVEN_TRUE" : "FALSE";
} else {
    $allEven = true;
    foreach ($nums as $n) { if ($n % 2 !== 0) { $allEven = false; break; } }
    echo $allEven ? "ALL_EVEN_TRUE" : "FALSE";
}

__vybe_check(ob_get_clean(), "ALL_EVEN_TRUE");
