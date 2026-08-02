<?php
// vybe-test: php/variable_variables/variable_variable_assign_then_read
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

$key = 'color';
$$key = 'blue';
echo $$key;
