<?php
// vybe-test: php/variable_variables/variable_function_basic
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

function greet(string $name): string { return "Hi, $name!"; }
$fn = 'greet';
echo $fn('Alice');
