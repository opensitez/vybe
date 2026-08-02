<?php
// vybe-test: php/php_constants/constant_function_by_name
// origin: languages/php/tests/php/test_php_constants.rs
// vybe-test-mode: compile

define('TIMEOUT', 30);
$name = 'TIMEOUT';
echo constant($name);
