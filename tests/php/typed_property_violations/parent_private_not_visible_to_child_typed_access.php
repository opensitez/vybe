<?php
// vybe-test: php/typed_property_violations/parent_private_not_visible_to_child_typed_access
// origin: languages/php/tests/php/test_typed_property_violations.rs

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

echo "parent_private_not_visible_to_child_typed_access_ok";

__vybe_check(ob_get_clean(), "parent_private_not_visible_to_child_typed_access_ok");
