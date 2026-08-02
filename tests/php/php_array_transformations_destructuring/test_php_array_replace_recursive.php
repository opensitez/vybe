<?php
// vybe-test: php/php_array_transformations_destructuring/test_php_array_replace_recursive
// origin: languages/php/tests/php/test_php_array_transformations_destructuring.rs
// vybe-test-mode: compile

$base = ["citrus" => ["orange"], "berries" => ["blackberry"]];
$replacement = ["citrus" => ["pineapple"], "berries" => ["strawberry"]];
$basket = array_replace_recursive($base, $replacement);
print_r($basket);
