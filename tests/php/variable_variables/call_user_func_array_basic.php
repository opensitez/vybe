<?php
// vybe-test: php/variable_variables/call_user_func_array_basic
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

function sum3(int $a, int $b, int $c): int { return $a + $b + $c; }
echo call_user_func_array('sum3', [1, 2, 3]);
