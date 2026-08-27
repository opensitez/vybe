<?php
// vybe-test: php/magic_method_errors/invoke_with_correct_args
// origin: languages/php/tests/php/test_magic_method_errors.rs

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

echo "invoke_with_correct_args_ok";

__vybe_check(ob_get_clean(), "invoke_with_correct_args_ok");
