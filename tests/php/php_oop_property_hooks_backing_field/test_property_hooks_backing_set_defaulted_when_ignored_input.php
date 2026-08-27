<?php
// vybe-test: php/php_oop_property_hooks_backing_field/test_property_hooks_backing_set_defaulted_when_ignored_input
// origin: languages/php/tests/php/test_php_oop_property_hooks_backing_field.rs

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

echo "test_property_hooks_backing_set_defaulted_when_ignored_input_ok";

__vybe_check(ob_get_clean(), "test_property_hooks_backing_set_defaulted_when_ignored_input_ok");
