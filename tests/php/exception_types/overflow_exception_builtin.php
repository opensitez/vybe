<?php
// vybe-test: php/exception_types/overflow_exception_builtin
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

try {
    throw new OverflowException('stack overflow');
} catch (OverflowException $e) {
    echo $e->getMessage();
}
