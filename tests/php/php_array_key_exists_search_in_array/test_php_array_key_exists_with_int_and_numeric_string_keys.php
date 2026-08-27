<?php
// vybe-test: php/php_array_key_exists_search_in_array/test_php_array_key_exists_with_int_and_numeric_string_keys
// origin: languages/php/tests/php/test_php_array_key_exists_search_in_array.rs

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

echo "test_php_array_key_exists_with_int_and_numeric_string_keys_ok";

__vybe_check(ob_get_clean(), "test_php_array_key_exists_with_int_and_numeric_string_keys_ok");
