<?php
// vybe-test: php/variable_variables/call_user_func_array_spread
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

$args = [10, 20, 30];
$fn = fn(int ...$nums) => array_sum($nums);
echo call_user_func_array($fn, $args);
