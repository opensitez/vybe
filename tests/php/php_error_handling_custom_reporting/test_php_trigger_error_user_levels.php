<?php
// vybe-test: php/php_error_handling_custom_reporting/test_php_trigger_error_user_levels
// origin: languages/php/tests/php/test_php_error_handling_custom_reporting.rs

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

echo "test_php_trigger_error_user_levels_ok";

__vybe_check(ob_get_clean(), "test_php_trigger_error_user_levels_ok");
