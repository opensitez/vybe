<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_object_property_storing_static_callable
// origin: languages/php/tests/php/test_php_dynamic_calling.rs

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

echo "php_dynamic_calling_object_property_storing_static_callable_ok";

__vybe_check(ob_get_clean(), "php_dynamic_calling_object_property_storing_static_callable_ok");
