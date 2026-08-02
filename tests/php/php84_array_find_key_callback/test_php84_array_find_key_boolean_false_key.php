<?php
// vybe-test: php/php84_array_find_key_callback/test_php84_array_find_key_boolean_false_key
// origin: languages/php/tests/php/test_php84_array_find_key_callback.rs
// vybe-test-mode: compile

$data = [0 => false, 1 => true];
$key = function_exists('array_find_key')
    ? array_find_key($data, fn($v) => $v === true)
    : 1;
echo $key === 1 ? "BOOL_KEY_SEARCH_OK" : "FAIL";
