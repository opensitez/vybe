<?php
// vybe-test: php/scope_variables/constant_in_namespace
// origin: languages/php/tests/php/test_scope_variables.rs

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

echo "constant_in_namespace_ok";

__vybe_check(ob_get_clean(), "constant_in_namespace_ok");
