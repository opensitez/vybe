<?php
// vybe-test: php/function_builtins/function_exists_user_defined
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

echo function_exists('myFunc') ? 'yes' : 'no';
function myFunc() {}
echo function_exists('myFunc') ? 'yes' : 'no';
