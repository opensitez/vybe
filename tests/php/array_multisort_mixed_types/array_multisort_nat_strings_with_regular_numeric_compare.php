<?php
// vybe-test: php/array_multisort_mixed_types/array_multisort_nat_strings_with_regular_numeric_compare
// origin: languages/php/tests/php/test_array_multisort_mixed_types.rs

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

echo "array_multisort_nat_strings_with_regular_numeric_compare_ok";

__vybe_check(ob_get_clean(), "array_multisort_nat_strings_with_regular_numeric_compare_ok");
