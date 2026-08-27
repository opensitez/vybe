<?php
// vybe-test: php/spl_structure_errors/spl_object_storage_detach_unknown_object_is_noop
// origin: languages/php/tests/php/test_spl_structure_errors.rs

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

echo "spl_object_storage_detach_unknown_object_is_noop_ok";

__vybe_check(ob_get_clean(), "spl_object_storage_detach_unknown_object_is_noop_ok");
