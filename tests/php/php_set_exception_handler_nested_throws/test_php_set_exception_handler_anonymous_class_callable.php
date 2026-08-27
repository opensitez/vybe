<?php
// vybe-test: php/php_set_exception_handler_nested_throws/test_php_set_exception_handler_anonymous_class_callable
// origin: languages/php/tests/php/test_php_set_exception_handler_nested_throws.rs

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

echo "test_php_set_exception_handler_anonymous_class_callable_ok";

__vybe_check(ob_get_clean(), "test_php_set_exception_handler_anonymous_class_callable_ok");
