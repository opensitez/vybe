<?php
// vybe-test: php/error_reporting_runtime/error_log_returns_true
// origin: languages/php/tests/php/test_error_reporting_runtime.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo error_log('test', 3, sys_get_temp_dir() . '/vybe_php_test.log') ? '1' : '0';

__vybe_check(ob_get_clean(), "1");
