<?php
// vybe-test: php/php_datetime_create_from_format_errors/test_datetime_create_from_format_trailing_space
// origin: languages/php/tests/php/test_php_datetime_create_from_format_errors.rs

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

echo "test_datetime_create_from_format_trailing_space_ok";

__vybe_check(ob_get_clean(), "test_datetime_create_from_format_trailing_space_ok");
