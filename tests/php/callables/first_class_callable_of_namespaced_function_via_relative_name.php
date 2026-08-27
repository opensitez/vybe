<?php
// vybe-test: php/callables/first_class_callable_of_namespaced_function_via_relative_name
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

echo "first_class_callable_of_namespaced_function_via_relative_name_ok";

__vybe_check(ob_get_clean(), "first_class_callable_of_namespaced_function_via_relative_name_ok");
