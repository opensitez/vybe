<?php
// vybe-test: php/magic_method_errors/readonly_dynamic_property_blocked
// origin: languages/php/tests/php/test_magic_method_errors.rs

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

echo "readonly_dynamic_property_blocked_ok";

__vybe_check(ob_get_clean(), "readonly_dynamic_property_blocked_ok");
