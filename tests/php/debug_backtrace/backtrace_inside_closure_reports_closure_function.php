<?php
// vybe-test: php/debug_backtrace/backtrace_inside_closure_reports_closure_function
// origin: languages/php/tests/php/test_debug_backtrace.rs

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

echo "backtrace_inside_closure_reports_closure_function_ok";

__vybe_check(ob_get_clean(), "backtrace_inside_closure_reports_closure_function_ok");
