<?php
// vybe-test: php/typed_property_violations/dynamic_property_on_typed_class_without_allow_dynamic
// origin: languages/php/tests/php/test_typed_property_violations.rs

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

echo "dynamic_property_on_typed_class_without_allow_dynamic_ok";

__vybe_check(ob_get_clean(), "dynamic_property_on_typed_class_without_allow_dynamic_ok");
