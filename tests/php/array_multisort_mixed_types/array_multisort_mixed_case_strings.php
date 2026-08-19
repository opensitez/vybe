<?php
// vybe-test: php/array_multisort_mixed_types/array_multisort_mixed_case_strings
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

$keys = ["beta", "Alpha", "charlie", "Bravo"];
$names = ["B", "A", "C", "D"];
array_multisort($keys, SORT_ASC, SORT_STRING | SORT_FLAG_CASE, $names, SORT_ASC, SORT_STRING);
echo implode(',', $keys) . "|" . implode(',', $names);

__vybe_check(ob_get_clean(), "Alpha,beta,Bravo,charlie|A,B,D,C");
