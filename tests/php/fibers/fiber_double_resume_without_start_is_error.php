<?php
// vybe-test: php/fibers/fiber_double_resume_without_start_is_error
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

echo "fiber_double_resume_without_start_is_error_ok";

__vybe_check(ob_get_clean(), "fiber_double_resume_without_start_is_error_ok");
