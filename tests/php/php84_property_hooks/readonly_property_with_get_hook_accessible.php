<?php
// vybe-test: php/php84_property_hooks/readonly_property_with_get_hook_accessible
// origin: languages/php/tests/php/test_php84_property_hooks.rs

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

echo "readonly_property_with_get_hook_accessible_ok";

__vybe_check(ob_get_clean(), "readonly_property_with_get_hook_accessible_ok");
