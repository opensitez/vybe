<?php
// vybe-test: php/function_builtins/call_user_func_closure
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

$double = function(int $n): int { return $n * 2; };
echo call_user_func($double, 21);
