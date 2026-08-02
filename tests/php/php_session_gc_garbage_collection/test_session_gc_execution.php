<?php
// vybe-test: php/php_session_gc_garbage_collection/test_session_gc_execution
// origin: languages/php/tests/php/test_php_session_gc_garbage_collection.rs

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

if (function_exists('session_gc')) {
    $deleted = session_gc();
    echo (is_int($deleted) || $deleted === false) ? 'session_gc_ok' : 'err', "\n";
} else {
    echo "session_gc_ok\n";
}

__vybe_check(ob_get_clean(), "session_gc_ok");
