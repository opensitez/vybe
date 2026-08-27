<?php
// vybe-test: php/math_advanced/modf_integer_and_fraction
// origin: languages/php/tests/php/test_math_advanced.rs

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

echo "modf_integer_and_fraction_ok";

__vybe_check(ob_get_clean(), "modf_integer_and_fraction_ok");
