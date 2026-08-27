<?php
// vybe-test: php/magic_constants_runtime/namespace_magic_inside_namespace
// origin: languages/php/tests/php/test_magic_constants_runtime.rs

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

echo "namespace_magic_inside_namespace_ok";

__vybe_check(ob_get_clean(), "namespace_magic_inside_namespace_ok");
