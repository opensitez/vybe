<?php
// vybe-test: php/variable_functions/get_defined_vars_returns_array
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

$alpha = 1;
$beta  = 2;
$vars = get_defined_vars();
echo is_array($vars) ? 'yes' : 'no';
