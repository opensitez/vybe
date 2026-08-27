<?php
// vybe-test: php/compression_runtime/gzuncompress_false_on_invalid_data
// origin: languages/php/tests/php/test_compression_runtime.rs

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

echo "gzuncompress_false_on_invalid_data_ok";

__vybe_check(ob_get_clean(), "gzuncompress_false_on_invalid_data_ok");
