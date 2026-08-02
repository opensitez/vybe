<?php
// vybe-test: php/restore_exception_handler/restore_exception_handler_basic
// origin: languages/php/tests/php/test_restore_exception_handler.rs

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

set_exception_handler(function($e) { echo "A"; });
set_exception_handler(function($e) { echo "B"; });
restore_exception_handler();

$old = set_exception_handler(function($e) { echo "C"; });
// We just check if it's callable without throwing
echo "ok";

__vybe_check(ob_get_clean(), "ok");
