<?php
// vybe-test: php/spl_runtime/spl_object_storage_count
// origin: languages/php/tests/php/test_spl_runtime.rs

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

echo "spl_object_storage_count_ok";

__vybe_check(ob_get_clean(), "spl_object_storage_count_ok");
