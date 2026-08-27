<?php
// vybe-test: php/readonly_class_php82/readonly_nested_object_reference_still_readonly_on_reassignment
// origin: languages/php/tests/php/test_readonly_class_php82.rs

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

echo "readonly_nested_object_reference_still_readonly_on_reassignment_ok";

__vybe_check(ob_get_clean(), "readonly_nested_object_reference_still_readonly_on_reassignment_ok");
