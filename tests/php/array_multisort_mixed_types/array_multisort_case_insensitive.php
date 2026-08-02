<?php
// vybe-test: php/array_multisort_mixed_types/array_multisort_case_insensitive
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

$arr = ["Alpha", "atomic", "Beta", "bank"];
array_multisort($arr, SORT_ASC, SORT_FLAG_CASE | SORT_STRING);
echo implode(',', $arr);

__vybe_check(ob_get_clean(), "Alpha,atomic,bank,Beta");
