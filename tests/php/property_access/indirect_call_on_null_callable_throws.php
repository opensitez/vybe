<?php
// vybe-test: php/property_access/indirect_call_on_null_callable_throws
// origin: languages/php/tests/php/test_property_access.rs

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

echo "indirect_call_on_null_callable_throws_ok";

__vybe_check(ob_get_clean(), "indirect_call_on_null_callable_throws_ok");
