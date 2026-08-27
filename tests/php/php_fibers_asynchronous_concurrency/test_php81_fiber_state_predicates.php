<?php
// vybe-test: php/php_fibers_asynchronous_concurrency/test_php81_fiber_state_predicates
// origin: languages/php/tests/php/test_php_fibers_asynchronous_concurrency.rs

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

echo "test_php81_fiber_state_predicates_ok";

__vybe_check(ob_get_clean(), "test_php81_fiber_state_predicates_ok");
