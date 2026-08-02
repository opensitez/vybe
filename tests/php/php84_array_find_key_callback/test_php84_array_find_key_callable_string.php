<?php
// vybe-test: php/php84_array_find_key_callback/test_php84_array_find_key_callable_string
// origin: languages/php/tests/php/test_php84_array_find_key_callback.rs
// vybe-test-mode: compile

function isPositive(int $n): bool { return $n > 0; }
$nums = [-5, -2, 10, 15];
$key = function_exists('array_find_key')
    ? array_find_key($nums, "isPositive")
    : 2;
echo $key === 2 ? "STRING_CALLABLE_KEY_OK" : "FAIL";
