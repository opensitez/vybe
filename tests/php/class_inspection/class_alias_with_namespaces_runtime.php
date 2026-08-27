<?php
// vybe-test: php/class_inspection/class_alias_with_namespaces_runtime
// origin: languages/php/tests/php/test_class_inspection.rs

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

echo "class_alias_with_namespaces_runtime_ok";

__vybe_check(ob_get_clean(), "class_alias_with_namespaces_runtime_ok");
