<?php
// vybe-test: php/error_handler/handler_return_false_leaves_error_get_last_populated
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

set_error_handler(fn() => false);
trigger_error('persist', E_USER_NOTICE);
restore_error_handler();
$e = error_get_last();
echo $e !== null && str_contains($e['message'], 'persist') ? 'last' : 'none';

__vybe_check(ob_get_clean(), "last");
