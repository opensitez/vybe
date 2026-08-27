<?php
// vybe-test: php/array_callbacks/array_filter_use_both_value_and_key
// origin: languages/php/tests/php/test_array_callbacks.rs

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

echo "array_filter_use_both_value_and_key_ok";

__vybe_check(ob_get_clean(), "array_filter_use_both_value_and_key_ok");
