<?php
// vybe-test: php/php_set_exception_handler_nested_throws/test_php_set_exception_handler_null_resets
// origin: languages/php/tests/php/test_php_set_exception_handler_nested_throws.rs
// vybe-test-mode: compile

set_exception_handler(fn($e) => null);
set_exception_handler(null);
echo "NULL_RESET_EXCEPTION_HANDLER_OK";
