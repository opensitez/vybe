<?php
// vybe-test: php/exception_types/exception_get_message_code_line
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

try {
    throw new Exception('test message', 42);
} catch (Exception $e) {
    echo $e->getMessage();
    echo $e->getCode();
    echo $e->getLine();
}
