<?php
// vybe-test: php/fiber_lifecycle/fiber_continues_after_catching_thrown_exception
// origin: languages/php/tests/php/test_fiber_lifecycle.rs

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

echo "fiber_continues_after_catching_thrown_exception_ok";

__vybe_check(ob_get_clean(), "fiber_continues_after_catching_thrown_exception_ok");
