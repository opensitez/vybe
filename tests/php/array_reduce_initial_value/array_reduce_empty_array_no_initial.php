<?php
// vybe-test: php/array_reduce_initial_value/array_reduce_empty_array_no_initial
// origin: languages/php/tests/php/test_array_reduce_initial_value.rs

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

$res = array_reduce([], function($c, $i) { return $c + $i; });
echo is_null($res) ? "null" : "not null";

__vybe_check(ob_get_clean(), "null");
