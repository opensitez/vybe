<?php
// vybe-test: php/php84_array_find_callback/test_php84_array_find_empty_array_returns_null
// origin: languages/php/tests/php/test_php84_array_find_callback.rs
// vybe-test-mode: compile

$res = function_exists('array_find')
    ? array_find([], fn($x) => true)
    : null;
echo $res === null ? "EMPTY_NULL_OK" : "FAIL";
