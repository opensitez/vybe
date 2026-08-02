<?php
// vybe-test: php/php84_array_find_key_callback/test_php84_array_find_key_zero_key
// origin: languages/php/tests/php/test_php84_array_find_key_callback.rs
// vybe-test-mode: compile

$arr = [0 => "zero_val", 1 => "one_val"];
$key = function_exists('array_find_key')
    ? array_find_key($arr, fn($v) => $v === "zero_val")
    : 0;
echo $key === 0 ? "ZERO_KEY_OK" : "FAIL";
