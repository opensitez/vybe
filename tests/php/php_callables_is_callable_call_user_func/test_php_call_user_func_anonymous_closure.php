<?php
// vybe-test: php/php_callables_is_callable_call_user_func/test_php_call_user_func_anonymous_closure
// origin: languages/php/tests/php/test_php_callables_is_callable_call_user_func.rs
// vybe-test-mode: compile

$closure = function($a, $b) { return $a - $b; };
echo call_user_func($closure, 100, 30);
