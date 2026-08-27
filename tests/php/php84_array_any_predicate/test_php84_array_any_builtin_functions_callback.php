<?php
// vybe-test: php/php84_array_any_predicate/test_php84_array_any_builtin_functions_callback
// origin: languages/php/tests/php/test_php84_array_any_predicate.rs

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

echo "test_php84_array_any_builtin_functions_callback_ok";

__vybe_check(ob_get_clean(), "test_php84_array_any_builtin_functions_callback_ok");
