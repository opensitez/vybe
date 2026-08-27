<?php
// vybe-test: php/php_functions_arrow_fn_variadic_named/test_php_named_arguments_with_variadic_args
// origin: languages/php/tests/php/test_php_functions_arrow_fn_variadic_named.rs

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

echo "test_php_named_arguments_with_variadic_args_ok";

__vybe_check(ob_get_clean(), "test_php_named_arguments_with_variadic_args_ok");
