<?php
// vybe-test: php/php_array_transformations_destructuring/test_php_array_merge_recursive_semantics
// origin: languages/php/tests/php/test_php_array_transformations_destructuring.rs
// vybe-test-mode: compile

$ar1 = ["color" => ["favorite" => "red"], 5];
$ar2 = [10, "color" => ["favorite" => "green", "blue"]];
$result = array_merge_recursive($ar1, $ar2);
print_r($result);
