<?php
// vybe-test: php/php_array_is_list_validation/test_array_is_list_string_numeric_keys_are_not_list
// origin: languages/php/tests/php/test_php_array_is_list_validation.rs

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

echo "test_array_is_list_string_numeric_keys_are_not_list_ok";

__vybe_check(ob_get_clean(), "test_array_is_list_string_numeric_keys_are_not_list_ok");
