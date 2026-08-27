<?php
// vybe-test: php/php_streams_file_system_wrapper_ops/test_php_stream_wrapper_register_custom
// origin: languages/php/tests/php/test_php_streams_file_system_wrapper_ops.rs

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

echo "test_php_stream_wrapper_register_custom_ok";

__vybe_check(ob_get_clean(), "test_php_stream_wrapper_register_custom_ok");
