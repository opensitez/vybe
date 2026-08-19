<?php
// vybe-test: php/array_multisort_mixed_types/array_multisort_by_floats_and_strings
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

$vals = ["1.2", 3, 2.5, "2.5", 1];
$labels = ["a", "b", "c", "d", "e"];
array_multisort($vals, SORT_ASC, SORT_NUMERIC, $labels, SORT_ASC, SORT_STRING);
echo implode(',', $vals) . "|" . implode(',', $labels);

__vybe_check(ob_get_clean(), "1,1.2,2.5,2.5,3|e,a,c,d,b");
