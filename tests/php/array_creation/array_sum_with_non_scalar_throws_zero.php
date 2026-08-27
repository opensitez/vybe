<?php
// vybe-test: php/array_creation/array_sum_with_non_scalar_throws_zero
// origin: languages/php/tests/php/test_array_creation.rs

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

echo "array_sum_with_non_scalar_throws_zero_ok";

__vybe_check(ob_get_clean(), "array_sum_with_non_scalar_throws_zero_ok");
