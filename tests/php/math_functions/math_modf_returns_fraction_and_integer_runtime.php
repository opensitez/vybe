<?php
// vybe-test: php/math_functions/math_modf_returns_fraction_and_integer_runtime
// origin: languages/php/tests/php/test_math_functions.rs

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

echo "math_modf_returns_fraction_and_integer_runtime_ok";

__vybe_check(ob_get_clean(), "math_modf_returns_fraction_and_integer_runtime_ok");
