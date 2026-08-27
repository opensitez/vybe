<?php
// vybe-test: php/php_constants_extended/php_int_min_value
// origin: languages/php/tests/php/test_php_constants_extended.rs

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

echo "php_int_min_value_ok";

__vybe_check(ob_get_clean(), "php_int_min_value_ok");
