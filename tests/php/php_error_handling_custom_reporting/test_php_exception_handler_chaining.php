<?php
// vybe-test: php/php_error_handling_custom_reporting/test_php_exception_handler_chaining
// origin: languages/php/tests/php/test_php_error_handling_custom_reporting.rs
// vybe-test-mode: compile

$prev = set_exception_handler(fn($e) => echo "Handler 1");
$prev2 = set_exception_handler(fn($e) => echo "Handler 2");
restore_exception_handler();
