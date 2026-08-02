<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_array_multisort_preserves_parallel_rows
// origin: languages/php/tests/php/test_php_array_sorting_multisort_callbacks.rs

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

$name = ["a", "b", "c", "d"];
$scores = [20, 10, 20, 5];
$age = [30, 40, 25, 50];

array_multisort($scores, SORT_ASC, SORT_NUMERIC, $age, SORT_DESC, SORT_NUMERIC, $name);
echo "{$scores[0]}:{$age[0]}:{$name[0]}|{$scores[1]}:{$age[1]}:{$name[1]}|{$scores[2]}:{$age[2]}:{$name[2]}";

__vybe_check(ob_get_clean(), "5:50:d|10:40:b|20:30:a");
