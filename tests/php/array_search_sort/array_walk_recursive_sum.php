<?php
// vybe-test: php/array_search_sort/array_walk_recursive_sum
// origin: languages/php/tests/php/test_array_search_sort.rs

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

$a = ['x' => [1, 2], 'y' => [3]];
$s = 0;
array_walk_recursive($a, function ($v) use (&$s) { $s += $v; });
echo $s;

__vybe_check(ob_get_clean(), "6");
