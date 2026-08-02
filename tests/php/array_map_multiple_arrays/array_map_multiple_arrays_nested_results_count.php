<?php
// vybe-test: php/array_map_multiple_arrays/array_map_multiple_arrays_nested_results_count
// origin: languages/php/tests/php/test_array_map_multiple_arrays.rs

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

$nums = [1, 2, 3, 4];
$labels = ['a', 'b'];
$zipped = array_map(null, $nums, $labels);
echo count($zipped) . '|' . count($zipped[2]);

__vybe_check(ob_get_clean(), "4|2");
