<?php
// vybe-test: php/php_error_clear_last_reset/test_php_error_get_last_user_error_type
// origin: languages/php/tests/php/test_php_error_clear_last_reset.rs
// vybe-test-mode: compile

@trigger_error("User Error Level", E_USER_ERROR);
$err = error_get_last();
error_clear_last();
echo $err["type"] === E_USER_ERROR ? "E_USER_ERROR_MATCH_OK" : "FAIL";
