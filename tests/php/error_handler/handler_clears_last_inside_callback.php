<?php
// vybe-test: php/error_handler/handler_clears_last_inside_callback
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

set_error_handler(function() {
    error_clear_last();
    return true;
});
trigger_error('gone', E_USER_WARNING);
restore_error_handler();
echo error_get_last() === null ? 'cleared' : 'set';

__vybe_check(ob_get_clean(), "cleared");
