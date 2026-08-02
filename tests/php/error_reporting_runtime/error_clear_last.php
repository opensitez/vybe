<?php
// vybe-test: php/error_reporting_runtime/error_clear_last
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

set_error_handler(fn() => true);
trigger_error('x', E_USER_NOTICE);
restore_error_handler();
error_clear_last();
echo error_get_last() === null ? 'cleared' : 'set';

__vybe_check(ob_get_clean(), "cleared");
