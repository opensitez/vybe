<?php
// vybe-test: php/array_multisort_mixed_types/array_multisort_assoc_tie_breaker_preserves_payload_by_index
// origin: languages/php/tests/php/test_array_multisort_mixed_types.rs

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

$scores = [10, 10, 10, 20];
$ids = [5, 2, 9, 1];
$names = ["x", "y", "z", "w"];
array_multisort($scores, SORT_ASC, SORT_NUMERIC, $ids, SORT_ASC, SORT_NUMERIC, $names, SORT_ASC, SORT_STRING);
echo implode(',', $scores) . "|" . implode(',', $ids) . "|" . implode(',', $names);

__vybe_check(ob_get_clean(), "10,10,10,20|2,5,9,1|y,x,z,w");
