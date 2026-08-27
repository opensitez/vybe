<?php
// vybe-test: php/property_access/write_property_on_array_throws_type_error
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

echo "write_property_on_array_throws_type_error_ok";

__vybe_check(ob_get_clean(), "write_property_on_array_throws_type_error_ok");
