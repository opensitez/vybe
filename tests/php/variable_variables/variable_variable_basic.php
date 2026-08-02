<?php
// vybe-test: php/variable_variables/variable_variable_basic
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

$varName = 'greeting';
$$varName = 'Hello';
echo $greeting;
