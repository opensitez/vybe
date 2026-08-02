<?php
// vybe-test: php/variable_functions/list_swap_variables
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

$a = 'first';
$b = 'second';
[$a, $b] = [$b, $a];
echo $a;
echo $b;
