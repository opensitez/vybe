<?php
// vybe-test: php/php84_array_find_callback/test_php84_array_find_returns_first_matching_element
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

$items = [10, 25, 40, 55];
if (function_exists('array_find')) {
    $firstEvenOver20 = array_find($items, fn($val) => $val > 20 && $val % 2 === 0);
    echo "Found: $firstEvenOver20";
} else {
    // Polyfill fallback verification for PHP 8.4 semantics
    $firstEvenOver20 = null;
    foreach ($items as $val) {
        if ($val > 20 && $val % 2 === 0) { $firstEvenOver20 = $val; break; }
    }
    echo "Found: $firstEvenOver20";
}

__vybe_check(ob_get_clean(), "Found: 40");
