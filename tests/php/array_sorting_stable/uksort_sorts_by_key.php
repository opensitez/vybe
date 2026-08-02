<?php
// vybe-test: php/array_sorting_stable/uksort_sorts_by_key
// origin: languages/php/tests/php/test_array_sorting_stable.rs

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

$a = ['banana' => 2, 'apple' => 1, 'cherry' => 3];
uksort($a, fn($a,$b) => strcmp($a,$b));
echo implode(',', array_keys($a));

__vybe_check(ob_get_clean(), "apple,banana,cherry");
