<?php
// vybe-test: php/callables/call_user_func_array_with_named_parameters
// origin: languages/php/tests/php/test_callables.rs

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

echo "call_user_func_array_with_named_parameters_ok";

__vybe_check(ob_get_clean(), "call_user_func_array_with_named_parameters_ok");
