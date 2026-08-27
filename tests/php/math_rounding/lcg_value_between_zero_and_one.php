<?php
// vybe-test: php/math_rounding/lcg_value_between_zero_and_one
// origin: languages/php/tests/php/test_math_rounding.rs

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

echo "lcg_value_between_zero_and_one_ok";

__vybe_check(ob_get_clean(), "lcg_value_between_zero_and_one_ok");
