<?php
// vybe-test: php/php_error_clear_last_reset/test_php_error_get_last_returns_null_after_clear
// origin: languages/php/tests/php/test_php_error_clear_last_reset.rs
// vybe-test-mode: compile

@trigger_error("Test", E_USER_NOTICE);
error_clear_last();
echo error_get_last() === null ? "CLEAR_NULL_OK" : "FAIL";
