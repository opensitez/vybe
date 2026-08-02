<?php
// vybe-test: php/php_callables_is_callable_call_user_func/test_php_dynamic_function_name_invocation
// origin: languages/php/tests/php/test_php_callables_is_callable_call_user_func.rs
// vybe-test-mode: compile

$fnName = "strtoupper";
echo $fnName("dynamic function call");
