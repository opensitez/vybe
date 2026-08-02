<?php
// vybe-test: php/php_exceptions_multi_catch_throw_expr/test_php_multi_catch_exception_handling
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

class InvalidUserException extends Exception {}
class DatabaseException extends Exception {}

function process($type) {
    try {
        if ($type === 1) throw new InvalidUserException("User invalid");
        if ($type === 2) throw new DatabaseException("DB error");
    } catch (InvalidUserException | DatabaseException $e) {
        echo "CAUGHT: " . $e->getMessage();
    }
}

process(2);

__vybe_check(ob_get_clean(), "CAUGHT: DB error");
