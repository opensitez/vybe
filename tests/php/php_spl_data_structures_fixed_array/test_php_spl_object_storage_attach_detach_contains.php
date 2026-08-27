<?php
// vybe-test: php/php_spl_data_structures_fixed_array/test_php_spl_object_storage_attach_detach_contains
// origin: languages/php/tests/php/test_php_spl_data_structures_fixed_array.rs

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

echo "test_php_spl_object_storage_attach_detach_contains_ok";

__vybe_check(ob_get_clean(), "test_php_spl_object_storage_attach_detach_contains_ok");
