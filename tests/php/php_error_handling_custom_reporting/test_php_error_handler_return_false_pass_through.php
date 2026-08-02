<?php
// vybe-test: php/php_error_handling_custom_reporting/test_php_error_handler_return_false_pass_through
// origin: languages/php/tests/php/test_php_error_handling_custom_reporting.rs
// vybe-test-mode: compile

set_error_handler(function($errno, $errstr) {
    return false; // pass to default PHP error handler
});
@trigger_error("Pass through notice", E_USER_NOTICE);
restore_error_handler();
