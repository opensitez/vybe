<?php
// vybe-test: php/php_string_manipulation/php_string_implode_handles_nested_arrays
// origin: languages/php/tests/php/test_php_string_manipulation.rs

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

echo "php_string_implode_handles_nested_arrays_ok";

__vybe_check(ob_get_clean(), "php_string_implode_handles_nested_arrays_ok");
