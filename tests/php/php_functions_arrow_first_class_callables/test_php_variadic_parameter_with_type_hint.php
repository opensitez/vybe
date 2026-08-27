<?php
// vybe-test: php/php_functions_arrow_first_class_callables/test_php_variadic_parameter_with_type_hint
// origin: languages/php/tests/php/test_php_functions_arrow_first_class_callables.rs

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

echo "test_php_variadic_parameter_with_type_hint_ok";

__vybe_check(ob_get_clean(), "test_php_variadic_parameter_with_type_hint_ok");
