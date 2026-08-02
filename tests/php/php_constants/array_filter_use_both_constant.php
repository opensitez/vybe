<?php
// vybe-test: php/php_constants/array_filter_use_both_constant
// origin: languages/php/tests/php/test_php_constants.rs
// vybe-test-mode: compile

$arr = ['x' => 10, 'y' => 5, 'z' => 20];
$result = array_filter($arr, fn($v, $k) => $k !== 'y' && $v > 8, ARRAY_FILTER_USE_BOTH);
echo count($result);
