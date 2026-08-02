<?php
// vybe-test: php/exception_types/exception_in_finally
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

try {
    try {
        throw new Exception('first');
    } finally {
        throw new Exception('from finally');
    }
} catch (Exception $e) {
    echo $e->getMessage();
}
