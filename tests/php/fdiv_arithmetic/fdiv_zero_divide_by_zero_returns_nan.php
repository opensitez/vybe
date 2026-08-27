<?php
// vybe-test: php/fdiv_arithmetic/fdiv_zero_divide_by_zero_returns_nan
// origin: languages/php/tests/php/test_fdiv_arithmetic.rs

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

echo "fdiv_zero_divide_by_zero_returns_nan_ok";

__vybe_check(ob_get_clean(), "fdiv_zero_divide_by_zero_returns_nan_ok");
