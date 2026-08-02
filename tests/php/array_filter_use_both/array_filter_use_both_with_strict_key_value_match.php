<?php
// vybe-test: php/array_filter_use_both/array_filter_use_both_with_strict_key_value_match
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

$arr = ['a' => 1, 'b' => 2, 'c' => '2'];
$res = array_filter($arr, fn($v, $k) => $k === 'a' || $v === 2, ARRAY_FILTER_USE_BOTH);
ksort($res);
echo implode('|', array_keys($res)) . '|' . implode(',', array_values($res));

__vybe_check(ob_get_clean(), "a,b|1,2");
