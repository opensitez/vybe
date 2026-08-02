<?php
// vybe-test: php/array_sorting_stable/array_multisort_with_same_values_and_payload_reordered_by_secondary
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

$scores = [1,1,1,2];
$labels = ['d','a','c','b'];
array_multisort($scores, SORT_ASC, SORT_NUMERIC, $labels, SORT_ASC, SORT_STRING);
echo implode(',', $labels);

__vybe_check(ob_get_clean(), "a,c,d,b");
