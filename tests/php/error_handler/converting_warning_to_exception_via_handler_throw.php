<?php
// vybe-test: php/error_handler/converting_warning_to_exception_via_handler_throw
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

set_error_handler(function($no, $msg) { throw new RuntimeException($msg); });
try {
    trigger_error('boom', E_USER_WARNING);
    echo 'no';
} catch (RuntimeException $e) {
    echo $e->getMessage();
} finally {
    restore_error_handler();
}

__vybe_check(ob_get_clean(), "boom");
