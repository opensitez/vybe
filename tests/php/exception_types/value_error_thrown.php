<?php
// vybe-test: php/exception_types/value_error_thrown
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

try {
    throw new ValueError('value out of acceptable range');
} catch (ValueError $e) {
    echo $e->getMessage();
}
