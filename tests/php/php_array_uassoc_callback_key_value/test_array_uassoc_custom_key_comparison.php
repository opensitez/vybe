<?php
// vybe-test: php/php_array_uassoc_callback_key_value/test_array_uassoc_custom_key_comparison
// origin: languages/php/tests/php/test_php_array_uassoc_callback_key_value.rs

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

echo "test_array_uassoc_custom_key_comparison_ok";

__vybe_check(ob_get_clean(), "test_array_uassoc_custom_key_comparison_ok");
