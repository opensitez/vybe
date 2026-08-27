<?php
// vybe-test: php/php_error_clear_last_reset/test_php_error_get_last_user_error_type
// origin: languages/php/tests/php/test_php_error_clear_last_reset.rs

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

echo "test_php_error_get_last_user_error_type_ok";

__vybe_check(ob_get_clean(), "test_php_error_get_last_user_error_type_ok");
