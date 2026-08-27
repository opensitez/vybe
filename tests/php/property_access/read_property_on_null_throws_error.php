<?php
// vybe-test: php/property_access/read_property_on_null_throws_error
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

echo "read_property_on_null_throws_error_ok";

__vybe_check(ob_get_clean(), "read_property_on_null_throws_error_ok");
