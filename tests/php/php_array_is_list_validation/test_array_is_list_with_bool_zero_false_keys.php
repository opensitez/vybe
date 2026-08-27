<?php
// vybe-test: php/php_array_is_list_validation/test_array_is_list_with_bool_zero_false_keys
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

echo "test_array_is_list_with_bool_zero_false_keys_ok";

__vybe_check(ob_get_clean(), "test_array_is_list_with_bool_zero_false_keys_ok");
