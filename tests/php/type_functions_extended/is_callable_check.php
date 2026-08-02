<?php
// vybe-test: php/type_functions_extended/is_callable_check
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

$fn = function($x) { return $x * 2; };
echo is_callable($fn) ? 'yes' : 'no';
echo is_callable(42) ? 'yes' : 'no';
