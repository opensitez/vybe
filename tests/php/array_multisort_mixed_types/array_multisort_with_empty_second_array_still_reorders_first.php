<?php
// vybe-test: php/array_multisort_mixed_types/array_multisort_with_empty_second_array_still_reorders_first
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

$scores = [8, 2, 5];
$payload = [];
array_multisort($scores, SORT_DESC, SORT_NUMERIC, $payload);
echo implode(',', $scores);

__vybe_check(ob_get_clean(), "8,5,2");
