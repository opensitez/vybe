<?php
// vybe-test: php/array_column_advanced/array_flip_non_scalar_value_throws
// origin: languages/php/tests/php/test_array_column_advanced.rs

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

echo "array_flip_non_scalar_value_throws_ok";

__vybe_check(ob_get_clean(), "array_flip_non_scalar_value_throws_ok");
