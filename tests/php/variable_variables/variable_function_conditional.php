<?php
// vybe-test: php/variable_variables/variable_function_conditional
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

function add(int $a, int $b): int { return $a + $b; }
function sub(int $a, int $b): int { return $a - $b; }
$mode = 'add';
$fn = $mode;
echo $fn(10, 3);
