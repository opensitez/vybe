<?php
// vybe-test: php/php84_array_find_callback/test_php84_array_find_truthy_zero_value
// origin: languages/php/tests/php/test_php84_array_find_callback.rs
// vybe-test-mode: compile

$nums = [-5, 0, 5];
$zero = function_exists('array_find')
    ? array_find($nums, fn($n) => $n === 0)
    : 0;
echo $zero === 0 ? "ZERO_MATCH_OK" : "FAIL";
