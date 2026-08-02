<?php
// vybe-test: php/error_handler/handler_prefixes_message_with_custom_tag
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

$captured = '';
set_error_handler(function($no, $msg) use (&$captured): bool {
    $captured = '[ERR]' . $msg;
    return true;
});
trigger_error('payload', E_USER_ERROR);
restore_error_handler();
echo $captured;

__vybe_check(ob_get_clean(), "[ERR]payload");
