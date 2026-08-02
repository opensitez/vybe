<?php
// vybe-test: php/exception_types/underflow_exception_builtin
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

try {
    throw new UnderflowException('stack underflow');
} catch (UnderflowException $e) {
    echo $e->getMessage();
}
