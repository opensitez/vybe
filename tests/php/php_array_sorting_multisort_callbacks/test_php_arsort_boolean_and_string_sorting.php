<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_arsort_boolean_and_string_sorting
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

$values = ["alpha" => "10", "beta" => "2", "gamma" => "A", "delta" => "9", "epsilon" => "0"];
arsort($values, SORT_NUMERIC);
echo implode("|", array_values($values));

__vybe_check(ob_get_clean(), "10|9|2|A|0");
