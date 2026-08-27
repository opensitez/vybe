<?php
// vybe-test: php/scope_variables/dynamic_class_name_with_namespace_scope
// origin: languages/php/tests/php/test_scope_variables.rs

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

echo "dynamic_class_name_with_namespace_scope_ok";

__vybe_check(ob_get_clean(), "dynamic_class_name_with_namespace_scope_ok");
