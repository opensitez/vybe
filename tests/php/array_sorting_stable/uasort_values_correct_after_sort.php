<?php
// vybe-test: php/array_sorting_stable/uasort_values_correct_after_sort
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

$a = ['x' => 30, 'y' => 10, 'z' => 20];
uasort($a, fn($a,$b) => $a <=> $b);
echo implode(',', $a);

__vybe_check(ob_get_clean(), "10,20,30");
