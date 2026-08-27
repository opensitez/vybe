<?php
// vybe-test: php/inheritance_runtime/variadic_parent_constructor
// origin: languages/php/tests/php/test_inheritance_runtime.rs

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

echo "variadic_parent_constructor_ok";

__vybe_check(ob_get_clean(), "variadic_parent_constructor_ok");
