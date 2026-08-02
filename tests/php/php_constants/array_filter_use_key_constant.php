<?php
// vybe-test: php/php_constants/array_filter_use_key_constant
// origin: languages/php/tests/php/test_php_constants.rs
// vybe-test-mode: compile

$arr = ['a' => 1, 'b' => 2, 'c' => 3];
$result = array_filter($arr, fn($k) => $k !== 'b', ARRAY_FILTER_USE_KEY);
echo count($result);
