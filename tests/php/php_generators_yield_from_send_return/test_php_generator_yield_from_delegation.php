<?php
// vybe-test: php/php_generators_yield_from_send_return/test_php_generator_yield_from_delegation
// origin: languages/php/tests/php/test_php_generators_yield_from_send_return.rs

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

echo "test_php_generator_yield_from_delegation_ok";

__vybe_check(ob_get_clean(), "test_php_generator_yield_from_delegation_ok");
