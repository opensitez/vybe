<?php
// vybe-test: php/spl_object_storage/spl_attach_with_data
// origin: languages/php/tests/php/test_spl_object_storage.rs

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

echo "spl_attach_with_data_ok";

__vybe_check(ob_get_clean(), "spl_attach_with_data_ok");
