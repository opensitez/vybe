<?php
// vybe-test: php/php_exceptions_multi_catch_throw_expr/test_php_exception_chaining_previous
// origin: languages/php/tests/php/test_php_exceptions_multi_catch_throw_expr.rs

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

try {
    try {
        throw new RuntimeException("Low level I/O fail", 101);
    } catch (RuntimeException $e) {
        throw new LogicException("High level processing error", 500, $e);
    }
} catch (LogicException $e) {
    echo $e->getMessage() . " -> " . $e->getPrevious()->getMessage();
}

__vybe_check(ob_get_clean(), "High level processing error -> Low level I/O fail");
