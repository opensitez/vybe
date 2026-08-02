<?php
// vybe-test: php/php_error_handling_custom_reporting/test_php_trigger_error_user_levels
// origin: languages/php/tests/php/test_php_error_handling_custom_reporting.rs
// vybe-test-mode: compile

@trigger_error("User error message", E_USER_ERROR);
@trigger_error("User notice message", E_USER_NOTICE);
@trigger_error("User deprecated message", E_USER_DEPRECATED);
