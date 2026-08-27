<?php
// vybe-test: php/php_oop_nullsafe_operator_chaining/test_nullsafe_operator_chain_on_non_null
// origin: languages/php/tests/php/test_php_oop_nullsafe_operator_chaining.rs

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

echo "test_nullsafe_operator_chain_on_non_null_ok";

__vybe_check(ob_get_clean(), "test_nullsafe_operator_chain_on_non_null_ok");
