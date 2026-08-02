<?php
// vybe-test: php/php_exceptions_rethrowing_previous/test_php_finally_block_runs_on_uncaught_exception
// origin: languages/php/tests/php/test_php_exceptions_rethrowing_previous.rs
// vybe-test-mode: compile

function testFinally() {
    try {
        throw new Exception("Fatal inside function");
    } finally {
        echo "FINALLY_EXECUTED ";
    }
}

try {
    testFinally();
} catch (Exception $e) {
    echo "CAUGHT_OUTSIDE";
}
