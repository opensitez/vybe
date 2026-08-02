<?php
// vybe-test: php/php84_array_find_key_callback/test_php84_array_find_key_empty_array
// origin: languages/php/tests/php/test_php84_array_find_key_callback.rs
// vybe-test-mode: compile

$key = function_exists('array_find_key')
    ? array_find_key([], fn($v) => true)
    : null;
echo $key === null ? "EMPTY_KEY_NULL_OK" : "FAIL";
