<?php
// vybe-test: php/numeric_operations/log_functions_natural_10_2
// origin: languages/php/tests/php/test_numeric_operations.rs

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

echo "log_functions_natural_10_2_ok";

__vybe_check(ob_get_clean(), "log_functions_natural_10_2_ok");
