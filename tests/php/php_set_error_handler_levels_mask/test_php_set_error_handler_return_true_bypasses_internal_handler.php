<?php
// vybe-test: php/php_set_error_handler_levels_mask/test_php_set_error_handler_return_true_bypasses_internal_handler
// origin: languages/php/tests/php/test_php_set_error_handler_levels_mask.rs

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

set_error_handler(function($errno, $errstr) {
    echo "CustomHandler: $errstr";
    return true; // Bypass standard PHP error handler
});

trigger_error("Bypassed error message", E_USER_NOTICE);
restore_error_handler();

__vybe_check(ob_get_clean(), "CustomHandler: Bypassed error message");
