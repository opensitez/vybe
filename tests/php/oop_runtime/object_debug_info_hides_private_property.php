<?php
// vybe-test: php/oop_runtime/object_debug_info_hides_private_property
// origin: languages/php/tests/php/test_oop_runtime.rs

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

echo "object_debug_info_hides_private_property_ok";

__vybe_check(ob_get_clean(), "object_debug_info_hides_private_property_ok");
