<?php
// vybe-test: php/function_builtins/call_user_func_array_basic
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

function add(int $a, int $b): int { return $a + $b; }
echo call_user_func_array('add', [10, 32]);
