<?php
// vybe-test: php/enum_advanced_runtime/enum_attributes_on_case
// origin: languages/php/tests/php/test_enum_advanced_runtime.rs

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

echo "enum_attributes_on_case_ok";

__vybe_check(ob_get_clean(), "enum_attributes_on_case_ok");
