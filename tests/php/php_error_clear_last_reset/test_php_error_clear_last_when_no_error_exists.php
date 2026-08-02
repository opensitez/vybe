<?php
// vybe-test: php/php_error_clear_last_reset/test_php_error_clear_last_when_no_error_exists
// origin: languages/php/tests/php/test_php_error_clear_last_reset.rs
// vybe-test-mode: compile

error_clear_last();
$err = error_get_last();
echo $err === null ? "NO_ERROR_NULL_OK" : "FAIL";
