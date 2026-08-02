<?php
// vybe-test: php/array_extra_builtins/array_rand_multiple_random_keys
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = ["a" => 1, "b" => 2, "c" => 3, "d" => 4, "e" => 5];
$keys = array_rand($a, 3);
echo count($keys);
echo is_array($keys) ? "array" : "not";
