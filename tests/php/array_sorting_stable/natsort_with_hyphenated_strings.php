<?php
// vybe-test: php/array_sorting_stable/natsort_with_hyphenated_strings
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

$labels = ['item-2a', 'item-10b', 'item-1c'];
natsort($labels);
echo implode(',', $labels);

__vybe_check(ob_get_clean(), "item-1c,item-2a,item-10b");
