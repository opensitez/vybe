<?php
// vybe-test: php/variable_functions/define_boolean_value
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

define('DEBUG_MODE', false);
define('FEATURE_FLAG', true);
echo DEBUG_MODE ? 'debug' : 'prod';
echo FEATURE_FLAG ? 'on' : 'off';
