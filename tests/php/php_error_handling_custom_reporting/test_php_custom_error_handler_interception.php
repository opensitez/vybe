<?php
// vybe-test: php/php_error_handling_custom_reporting/test_php_custom_error_handler_interception
// origin: languages/php/tests/php/test_php_error_handling_custom_reporting.rs

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

$captured = [];
set_error_handler(function($errno, $errstr) use (&$captured) {
    $captured[] = "ERR[$errno]: $errstr";
    return true; // suppress default PHP handler
});

trigger_error("Custom Warning Message", E_USER_WARNING);
restore_error_handler();

echo implode(", ", $captured);

__vybe_check(ob_get_clean(), "ERR[512]: Custom Warning Message");
