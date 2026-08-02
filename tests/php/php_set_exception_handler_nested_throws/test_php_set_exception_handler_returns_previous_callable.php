<?php
// vybe-test: php/php_set_exception_handler_nested_throws/test_php_set_exception_handler_returns_previous_callable
// origin: languages/php/tests/php/test_php_set_exception_handler_nested_throws.rs
// vybe-test-mode: compile

$first = fn($e) => null;
set_exception_handler($first);
$second = set_exception_handler(fn($e) => null);
restore_exception_handler();
restore_exception_handler();
echo $second === $first ? "PREVIOUS_HANDLER_RETURNED_OK" : "FAIL";
