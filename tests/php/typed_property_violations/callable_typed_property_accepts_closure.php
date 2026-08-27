<?php
// vybe-test: php/typed_property_violations/callable_typed_property_accepts_closure
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

echo "callable_typed_property_accepts_closure_ok";

__vybe_check(ob_get_clean(), "callable_typed_property_accepts_closure_ok");
