<?php
// vybe-test: php/error_handler/error_reporting_zero_suppresses_user_notice_to_handler
// origin: languages/php/tests/php/test_error_handler.rs

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

$fired = false;
set_error_handler(function() use (&$fired): bool { $fired = true; return true; });
$old = error_reporting(0);
trigger_error('hidden', E_USER_NOTICE);
error_reporting($old);
restore_error_handler();
echo $fired ? 'fired' : 'hidden';

__vybe_check(ob_get_clean(), "fired");
