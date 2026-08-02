<?php
// vybe-test: php/variable_functions/constant_get_value
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

define('MAX_RETRIES', 3);
$name = 'MAX_RETRIES';
echo constant($name);
