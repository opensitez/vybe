<?php
// vybe-test: php/output_buffering/ob_start_with_callable_function_name
// origin: languages/php/tests/php/test_output_buffering.rs

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

echo "ob_start_with_callable_function_name_ok";

__vybe_check(ob_get_clean(), "ob_start_with_callable_function_name_ok");
