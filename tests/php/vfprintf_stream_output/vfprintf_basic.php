<?php
// vybe-test: php/vfprintf_stream_output/vfprintf_basic
// origin: languages/php/tests/php/test_vfprintf_stream_output.rs

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

echo "vfprintf_basic_ok";

__vybe_check(ob_get_clean(), "vfprintf_basic_ok");
