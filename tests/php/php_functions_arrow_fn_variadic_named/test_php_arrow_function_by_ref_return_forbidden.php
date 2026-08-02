<?php
// vybe-test: php/php_functions_arrow_fn_variadic_named/test_php_arrow_function_by_ref_return_forbidden
// origin: languages/php/tests/php/test_php_functions_arrow_fn_variadic_named.rs
// vybe-test-mode: compile

$val = 100;
$getRef = fn&() => $val; // Arrow function return by reference syntax
