<?php
// vybe-test: php/php_error_clear_last_reset/test_php_error_get_last_deprecated_type
// origin: languages/php/tests/php/test_php_error_clear_last_reset.rs
// vybe-test-mode: compile

@trigger_error("Deprecation warning", E_USER_DEPRECATED);
$err = error_get_last();
error_clear_last();
echo $err["type"] === E_USER_DEPRECATED ? "E_USER_DEPRECATED_MATCH_OK" : "FAIL";
