<?php
// vybe-test: php/usort/array_multisort_sorts_parallel_arrays
// origin: languages/php/tests/php/test_usort.rs

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

$nums = [3, 1, 2];
$labels = ['c', 'a', 'b'];
array_multisort($nums, $labels);
echo implode('-', $nums) . ':' . implode('-', $labels);

__vybe_check(ob_get_clean(), "1-2-3:a-b-c");
