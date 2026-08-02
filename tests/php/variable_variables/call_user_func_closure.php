<?php
// vybe-test: php/variable_variables/call_user_func_closure
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

$mul = fn(int $a, int $b) => $a * $b;
echo call_user_func($mul, 6, 7);
