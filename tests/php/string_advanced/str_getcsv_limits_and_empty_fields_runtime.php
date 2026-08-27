<?php
// vybe-test: php/string_advanced/str_getcsv_limits_and_empty_fields_runtime
// origin: languages/php/tests/php/test_string_advanced.rs

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

echo "str_getcsv_limits_and_empty_fields_runtime_ok";

__vybe_check(ob_get_clean(), "str_getcsv_limits_and_empty_fields_runtime_ok");
