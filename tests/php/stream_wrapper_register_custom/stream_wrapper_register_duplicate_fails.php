<?php
// vybe-test: php/stream_wrapper_register_custom/stream_wrapper_register_duplicate_fails
// origin: languages/php/tests/php/test_stream_wrapper_register_custom.rs

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

echo "stream_wrapper_register_duplicate_fails_ok";

__vybe_check(ob_get_clean(), "stream_wrapper_register_duplicate_fails_ok");
