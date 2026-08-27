<?php
// vybe-test: php/callables/is_callable_on_array_with_string_instance
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

echo "is_callable_on_array_with_string_instance_ok";

__vybe_check(ob_get_clean(), "is_callable_on_array_with_string_instance_ok");
