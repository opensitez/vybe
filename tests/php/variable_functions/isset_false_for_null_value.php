<?php
// vybe-test: php/variable_functions/isset_false_for_null_value
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

$a = null;
$b = 0;
$c = '';
echo isset($a) ? 'set' : 'unset';
echo isset($b) ? 'set' : 'unset';
echo isset($c) ? 'set' : 'unset';
