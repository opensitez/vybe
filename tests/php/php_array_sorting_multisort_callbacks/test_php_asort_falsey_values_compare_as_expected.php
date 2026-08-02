<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_asort_falsey_values_compare_as_expected
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

$values = ["a" => 0, "b" => false, "c" => "00", "d" => 1];
asort($values, SORT_REGULAR);
echo implode("|", array_keys($values));

__vybe_check(ob_get_clean(), "a,b,c,d");
