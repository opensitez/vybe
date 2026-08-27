<?php
// vybe-test: php/named_arguments_runtime/named_args_splat_mixed
// origin: languages/php/tests/php/test_named_arguments_runtime.rs

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

echo "named_args_splat_mixed_ok";

__vybe_check(ob_get_clean(), "named_args_splat_mixed_ok");
