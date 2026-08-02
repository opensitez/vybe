<?php
// vybe-test: php/set_error_handler_bypass/set_error_handler_returns_false_bypasses
// origin: languages/php/tests/php/test_set_error_handler_bypass.rs

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

set_error_handler(function() {
    echo "caught|";
    return false; // Tells PHP to continue with normal error handler
});

@trigger_error("msg", E_USER_NOTICE);
echo "done";

__vybe_check(ob_get_clean(), "caught|done");
