<?php
// vybe-test: php/array_extra_builtins/array_rand_with_non_numeric_count_falls_back_to_scalar
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = ["a" => 1, "b" => 2, "c" => 3];
$key = array_rand($a, 2);
echo is_array($key) ? "array" : "scalar";
echo count($key);
