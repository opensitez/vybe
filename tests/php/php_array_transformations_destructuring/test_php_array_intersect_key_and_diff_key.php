<?php
// vybe-test: php/php_array_transformations_destructuring/test_php_array_intersect_key_and_diff_key
// origin: languages/php/tests/php/test_php_array_transformations_destructuring.rs
// vybe-test-mode: compile

$array1 = ['blue' => 1, 'red' => 2, 'green' => 3];
$array2 = ['green' => 5, 'yellow' => 7, 'cyan' => 8];
$intersect = array_intersect_key($array1, $array2);
$diff = array_diff_key($array1, $array2);
echo count($intersect) . count($diff);
