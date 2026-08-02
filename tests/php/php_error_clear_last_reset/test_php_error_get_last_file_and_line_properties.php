<?php
// vybe-test: php/php_error_clear_last_reset/test_php_error_get_last_file_and_line_properties
// origin: languages/php/tests/php/test_php_error_clear_last_reset.rs
// vybe-test-mode: compile

@trigger_error("File line test", E_USER_NOTICE);
$err = error_get_last();
error_clear_last();
echo isset($err["file"]) && isset($err["line"]) && $err["line"] > 0 ? "FILE_LINE_KEYS_OK" : "FAIL";
