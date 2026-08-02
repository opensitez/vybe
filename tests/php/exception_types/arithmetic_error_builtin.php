<?php
// vybe-test: php/exception_types/arithmetic_error_builtin
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

try {
    throw new ArithmeticError('division undefined');
} catch (ArithmeticError $e) {
    echo $e->getMessage();
}
