<?php
// vybe-test: php/array_multisort_mixed_types/array_multisort_large_and_edge_flags
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

$flags = [SORT_ASC, SORT_DESC];
$primary = ["a", "b", "c", "d"];
$secondary = [4, 3, 2, 1];
array_multisort($primary, $flags[0], SORT_STRING, $secondary, $flags[1], SORT_NUMERIC);
echo implode(',', $primary) . "|" . implode(',', $secondary);

__vybe_check(ob_get_clean(), "a,b,c,d|4,3,2,1");
