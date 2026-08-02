<?php
// vybe-test: php/php_set_error_handler_levels_mask/test_php_set_error_handler_receives_file_and_line
// origin: languages/php/tests/php/test_php_set_error_handler_levels_mask.rs
// vybe-test-mode: compile

$capturedLine = 0;
set_error_handler(function($errno, $errstr, $errfile, $errline) use (&$capturedLine) {
    $capturedLine = $errline;
    return true;
});
@trigger_error("Test line capture", E_USER_NOTICE);
restore_error_handler();
echo $capturedLine > 0 ? "LINE_CAPTURE_OK" : "FAIL";
