<?php
// vybe-test: php/covariant_return_types/child_overrides_with_own_class_return_type
// origin: languages/php/tests/php/test_covariant_return_types.rs

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

echo "child_overrides_with_own_class_return_type_ok";

__vybe_check(ob_get_clean(), "child_overrides_with_own_class_return_type_ok");
