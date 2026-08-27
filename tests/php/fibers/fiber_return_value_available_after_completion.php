<?php
// vybe-test: php/fibers/fiber_return_value_available_after_completion
// origin: languages/php/tests/php/test_fibers.rs

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

echo "fiber_return_value_available_after_completion_ok";

__vybe_check(ob_get_clean(), "fiber_return_value_available_after_completion_ok");
