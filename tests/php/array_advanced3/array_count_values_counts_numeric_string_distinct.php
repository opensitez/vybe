<?php
// vybe-test: php/array_advanced3/array_count_values_counts_numeric_string_distinct
// origin: languages/php/tests/php/test_array_advanced3.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "array_count_values_counts_numeric_string_distinct_ok";

__vybe_check(ob_get_clean(), "array_count_values_counts_numeric_string_distinct_ok");
