<?php
// vybe-test: php/php84_array_any_predicate/test_php84_array_any_builtin_functions_callback
// origin: languages/php/tests/php/test_php84_array_any_predicate.rs
// vybe-test-mode: compile

$items = ["123", "abc", "456"];
$hasNumeric = function_exists('array_any')
    ? array_any($items, "is_numeric")
    : true;
echo $hasNumeric ? "BUILTIN_CALLBACK_ANY_OK" : "FAIL";
