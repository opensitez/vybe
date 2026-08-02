<?php
// vybe-test: php/exception_types/logic_exception_builtin
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

try {
    throw new LogicException('precondition violated');
} catch (LogicException $e) {
    echo $e->getMessage();
}
