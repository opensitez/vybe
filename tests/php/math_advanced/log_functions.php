<?php
// vybe-test: php/math_advanced/log_functions
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

echo "log_functions_ok";

__vybe_check(ob_get_clean(), "log_functions_ok");
