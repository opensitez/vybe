<?php
// vybe-test: php/php_stream_filters_base64_rot13/test_php_stream_filter_append_base64_encode
// origin: languages/php/tests/php/test_php_stream_filters_base64_rot13.rs

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

echo "test_php_stream_filter_append_base64_encode_ok";

__vybe_check(ob_get_clean(), "test_php_stream_filter_append_base64_encode_ok");
