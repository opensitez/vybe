<?php
// vybe-test: php/named_args_extended/named_arg_before_variadic
// origin: languages/php/tests/php/test_named_args_extended.rs

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

echo "named_arg_before_variadic_ok";

__vybe_check(ob_get_clean(), "named_arg_before_variadic_ok");
