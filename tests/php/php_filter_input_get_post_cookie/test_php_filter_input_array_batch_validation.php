<?php
// vybe-test: php/php_filter_input_get_post_cookie/test_php_filter_input_array_batch_validation
// origin: languages/php/tests/php/test_php_filter_input_get_post_cookie.rs

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

echo "test_php_filter_input_array_batch_validation_ok";

__vybe_check(ob_get_clean(), "test_php_filter_input_array_batch_validation_ok");
