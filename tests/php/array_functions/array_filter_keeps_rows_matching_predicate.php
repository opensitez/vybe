<?php
// vybe-test: php/array_functions/array_filter_keeps_rows_matching_predicate
// origin: languages/php/tests/php/test_array_functions.rs

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

$rows = [['n' => 1], ['n' => 0], ['n' => 3]];
echo count(array_values(array_filter($rows, fn(array $r): bool => $r['n'] > 0)));

__vybe_check(ob_get_clean(), "2");
