<?php
// vybe-test: php/error_reporting_runtime/debug_backtrace_has_file_key
// origin: languages/php/tests/php/test_error_reporting_runtime.rs

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

echo "debug_backtrace_has_file_key_ok";

__vybe_check(ob_get_clean(), "debug_backtrace_has_file_key_ok");
