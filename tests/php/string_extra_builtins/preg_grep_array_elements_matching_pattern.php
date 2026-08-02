<?php
// vybe-test: php/string_extra_builtins/preg_grep_array_elements_matching_pattern
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$numbers = [1, 15, 3, 200, 42, 7, 100];
$large = preg_grep('/^[0-9]{3}/', array_map('strval', $numbers));
echo is_array($large) ? "array" : "not";
