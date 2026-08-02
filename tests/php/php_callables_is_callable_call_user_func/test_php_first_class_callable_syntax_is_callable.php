<?php
// vybe-test: php/php_callables_is_callable_call_user_func/test_php_first_class_callable_syntax_is_callable
// origin: languages/php/tests/php/test_php_callables_is_callable_call_user_func.rs
// vybe-test-mode: compile

$c = strlen(...);
echo is_callable($c) ? "IS_CALLABLE" : "FAIL";
