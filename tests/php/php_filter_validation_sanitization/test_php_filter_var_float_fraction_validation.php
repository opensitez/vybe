<?php
// vybe-test: php/php_filter_validation_sanitization/test_php_filter_var_float_fraction_validation
// origin: languages/php/tests/php/test_php_filter_validation_sanitization.rs
// vybe-test-mode: compile

$floatStr = "3.14159";
$res = filter_var($floatStr, FILTER_VALIDATE_FLOAT);
echo $res !== false ? "FLOAT_OK" : "FLOAT_FAIL";
