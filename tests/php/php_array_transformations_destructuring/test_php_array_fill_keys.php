<?php
// vybe-test: php/php_array_transformations_destructuring/test_php_array_fill_keys
// origin: languages/php/tests/php/test_php_array_transformations_destructuring.rs
// vybe-test-mode: compile

$keys = ["foo", 5, 10, "bar"];
$a = array_fill_keys($keys, "default");
print_r($a);
