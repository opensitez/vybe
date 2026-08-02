<?php
// vybe-test: php/php_filter_validation_sanitization/test_php_filter_var_null_on_failure_flag
// origin: languages/php/tests/php/test_php_filter_validation_sanitization.rs
// vybe-test-mode: compile

$res = filter_var("not_a_number", FILTER_VALIDATE_INT, FILTER_NULL_ON_FAILURE);
echo is_null($res) ? "NULL_ON_FAIL" : "OTHER";
