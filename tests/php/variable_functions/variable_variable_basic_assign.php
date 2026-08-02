<?php
// vybe-test: php/variable_functions/variable_variable_basic_assign
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

$varName = 'color';
$$varName = 'blue';
echo $color;
echo $$varName;
