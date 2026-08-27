<?php
// vybe-test: php/property_access/get_class_on_non_object_with_false_returns_false
// origin: languages/php/tests/php/test_property_access.rs

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

echo "get_class_on_non_object_with_false_returns_false_ok";

__vybe_check(ob_get_clean(), "get_class_on_non_object_with_false_returns_false_ok");
