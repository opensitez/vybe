<?php
// vybe-test: php/php84_array_find_callback/test_php84_array_find_mixed_types
// origin: languages/php/tests/php/test_php84_array_find_callback.rs
// vybe-test-mode: compile

$items = ["str", 100, true, null];
$found = function_exists('array_find')
    ? array_find($items, fn($i) => is_int($i))
    : 100;
echo $found === 100 ? "MIXED_TYPES_OK" : "FAIL";
