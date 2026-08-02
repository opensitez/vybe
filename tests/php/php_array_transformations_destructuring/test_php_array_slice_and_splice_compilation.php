<?php
// vybe-test: php/php_array_transformations_destructuring/test_php_array_slice_and_splice_compilation
// origin: languages/php/tests/php/test_php_array_transformations_destructuring.rs
// vybe-test-mode: compile

$input = ["red", "green", "blue", "yellow"];
$output = array_slice($input, 2);
$removed = array_splice($input, 1, 2, ["orange"]);
echo count($output) + count($removed);
