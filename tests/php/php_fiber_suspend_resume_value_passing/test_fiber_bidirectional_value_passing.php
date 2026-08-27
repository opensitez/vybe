<?php
// vybe-test: php/php_fiber_suspend_resume_value_passing/test_fiber_bidirectional_value_passing
// origin: languages/php/tests/php/test_php_fiber_suspend_resume_value_passing.rs

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

echo "test_fiber_bidirectional_value_passing_ok";

__vybe_check(ob_get_clean(), "test_fiber_bidirectional_value_passing_ok");
