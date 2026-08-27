<?php
// vybe-test: php/array_creation/range_negative_step_with_float
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

echo "range_negative_step_with_float_ok";

__vybe_check(ob_get_clean(), "range_negative_step_with_float_ok");
