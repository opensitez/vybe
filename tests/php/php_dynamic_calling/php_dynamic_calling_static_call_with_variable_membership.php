<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_static_call_with_variable_membership
// origin: languages/php/tests/php/test_php_dynamic_calling.rs

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

echo "php_dynamic_calling_static_call_with_variable_membership_ok";

__vybe_check(ob_get_clean(), "php_dynamic_calling_static_call_with_variable_membership_ok");
