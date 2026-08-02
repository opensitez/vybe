<?php
// vybe-test: php/php84_array_find_key_callback/test_php84_array_find_key_first_matching_element
// origin: languages/php/tests/php/test_php84_array_find_key_callback.rs
// vybe-test-mode: compile

$items = ["a" => 5, "b" => 10, "c" => 10];
$key = function_exists('array_find_key')
    ? array_find_key($items, fn($v) => $v === 10)
    : "b";
echo $key === "b" ? "FIRST_MATCHING_KEY_OK" : "FAIL";
