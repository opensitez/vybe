<?php
// vybe-test: php/php_exceptions_multi_catch_throw_expr/test_php_rethrowing_caught_exception
// origin: languages/php/tests/php/test_php_exceptions_multi_catch_throw_expr.rs
// vybe-test-mode: compile

function inner() {
    throw new Exception("Inner fail");
}

function outer() {
    try {
        inner();
    } catch (Exception $e) {
        // Log & rethrow
        throw $e;
    }
}

try {
    outer();
} catch (Exception $e) {
    echo $e->getMessage();
}
