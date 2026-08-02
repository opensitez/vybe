<?php
// vybe-test: php/php_callables_is_callable_call_user_func/test_php_is_callable_callable_name_out_param
// origin: languages/php/tests/php/test_php_callables_is_callable_call_user_func.rs
// vybe-test-mode: compile

$callableName = "";
$check = is_callable([stdClass::class, "nonExistent"], syntax_only: true, callable_name: $callableName);
echo "Check=$check Name=$callableName";
