<?php
// vybe-test: php/php_array_chunk_combine_count_values/test_php_array_count_values_string_number_coercion
// origin: languages/php/tests/php/test_php_array_chunk_combine_count_values.rs

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

$values = [1, "1", 2, "2", "2", 1];
$counts = array_count_values($values);
echo "1=" . $counts["1"] . " ";
echo "2=" . $counts["2"];

__vybe_check(ob_get_clean(), "1=3 2=3");
