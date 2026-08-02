<?php
// vybe-test: php/variable_functions/defined_check_constant
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

define('APP_VERSION', '1.0.0');
echo defined('APP_VERSION') ? 'yes' : 'no';
echo defined('MISSING_CONST') ? 'yes' : 'no';
