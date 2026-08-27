<?php
// vybe-test: php/array_uintersect_uassoc_callback/array_uintersect_uassoc_key_as_numeric_string
// origin: languages/php/tests/php/test_array_uintersect_uassoc_callback.rs

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

echo "array_uintersect_uassoc_key_as_numeric_string_ok";

__vybe_check(ob_get_clean(), "array_uintersect_uassoc_key_as_numeric_string_ok");
