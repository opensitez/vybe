<?php
// vybe-test: php/php_array_transformations_destructuring/test_php_array_filter_with_use_both_flag
// origin: languages/php/tests/php/test_php_array_transformations_destructuring.rs

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

$arr = ["a" => 1, "b" => 2, "c" => 3, "d" => 4];
$filtered = array_filter($arr, fn($val, $key) => $val > 2 && $key !== "d", ARRAY_FILTER_USE_BOTH);
echo implode(", ", array_keys($filtered));

__vybe_check(ob_get_clean(), "c");
