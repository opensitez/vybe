<?php
// vybe-test: php/php_error_clear_last_reset/test_php_error_clear_last_inside_error_handler
// origin: languages/php/tests/php/test_php_error_clear_last_reset.rs
// vybe-test-mode: compile

set_error_handler(function($errno, $errstr) {
    error_clear_last();
    return true;
});
@trigger_error("Inside handler", E_USER_NOTICE);
restore_error_handler();
echo error_get_last() === null ? "CLEAR_INSIDE_HANDLER_OK" : "FAIL";
