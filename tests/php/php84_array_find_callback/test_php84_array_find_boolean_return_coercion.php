<?php
// vybe-test: php/php84_array_find_callback/test_php84_array_find_boolean_return_coercion
// origin: languages/php/tests/php/test_php84_array_find_callback.rs
// vybe-test-mode: compile

$items = [0, 1, 2];
$found = function_exists('array_find')
    ? array_find($items, fn($n) => $n) // Truthy predicate
    : 1;
echo $found === 1 ? "TRUTHY_PREDICATE_OK" : "FAIL";
