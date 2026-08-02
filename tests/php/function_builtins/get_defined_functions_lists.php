<?php
// vybe-test: php/function_builtins/get_defined_functions_lists
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

function myCustomFn() {}
$all = get_defined_functions();
echo isset($all['user']) ? 'has user' : 'no user';
echo isset($all['internal']) ? ':has internal' : ':no internal';
echo in_array('mycustomfn', $all['user']) ? ':found' : ':not found';
