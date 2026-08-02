<?php
// vybe-test: php/variable_variables/call_user_func_basic
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

function double(int $n): int { return $n * 2; }
$result = call_user_func('double', 21);
echo $result;
