<?php
// vybe-test: php/php_error_clear_last_reset/test_php_error_clear_last_between_multiple_errors
// origin: languages/php/tests/php/test_php_error_clear_last_reset.rs
// vybe-test-mode: compile

@trigger_error("First notice", E_USER_NOTICE);
error_clear_last();
@trigger_error("Second notice", E_USER_NOTICE);
$err = error_get_last();
error_clear_last();
echo $err["message"] === "Second notice" ? "SECOND_NOTICE_CAPTURED_OK" : "FAIL";
