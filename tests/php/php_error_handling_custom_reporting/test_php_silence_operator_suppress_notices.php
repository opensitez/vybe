<?php
// vybe-test: php/php_error_handling_custom_reporting/test_php_silence_operator_suppress_notices
// origin: languages/php/tests/php/test_php_error_handling_custom_reporting.rs
// vybe-test-mode: compile

$arr = [];
$val = @$arr["non_existent_key"];
echo $val === null ? "NULL_SUPPRESSED" : "FAIL";
