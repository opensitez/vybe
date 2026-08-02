<?php
// vybe-test: php/array_filter_use_both/array_filter_with_lambda_returns_only_true_indexes
// origin: languages/php/tests/php/test_array_filter_use_both.rs

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

$arr = [0 => false, 1 => true, 2 => false, 3 => true];
$res = array_filter($arr, fn($v) => $v);
echo json_encode(array_values(array_keys($res)));

__vybe_check(ob_get_clean(), "[1,3]");
