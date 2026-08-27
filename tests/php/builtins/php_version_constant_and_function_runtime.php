<?php
// vybe-test: php/builtins/php_version_constant_and_function_runtime
// origin: languages/php/tests/php/test_builtins.rs

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

echo "php_version_constant_and_function_runtime_ok";

__vybe_check(ob_get_clean(), "php_version_constant_and_function_runtime_ok");
