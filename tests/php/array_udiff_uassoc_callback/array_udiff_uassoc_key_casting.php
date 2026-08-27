<?php
// vybe-test: php/array_udiff_uassoc_callback/array_udiff_uassoc_key_casting
// origin: languages/php/tests/php/test_array_udiff_uassoc_callback.rs

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

echo "array_udiff_uassoc_key_casting_ok";

__vybe_check(ob_get_clean(), "array_udiff_uassoc_key_casting_ok");
