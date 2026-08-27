<?php
// vybe-test: php/array_column_object_properties/array_column_objects_private_prop_access
// origin: languages/php/tests/php/test_array_column_object_properties.rs

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

echo "array_column_objects_private_prop_access_ok";

__vybe_check(ob_get_clean(), "array_column_objects_private_prop_access_ok");
