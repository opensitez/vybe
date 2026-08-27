<?php
// vybe-test: php/static_members_runtime/static_property_with_fully_qualified_reference
// origin: languages/php/tests/php/test_static_members_runtime.rs

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

echo "static_property_with_fully_qualified_reference_ok";

__vybe_check(ob_get_clean(), "static_property_with_fully_qualified_reference_ok");
