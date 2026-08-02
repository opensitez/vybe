<?php
// vybe-test: php/variable_variables/variable_variable_expression
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

$prefix = 'my';
$suffix = 'Var';
$varName = $prefix . $suffix;
$$varName = 42;
echo $myVar;
