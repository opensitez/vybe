<?php
// vybe-test: php/array_extra_builtins/array_rand_with_count_one_returns_scalar
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = [1, 2, 3, 4];
$key = array_rand($a, 1);
echo is_array($key) ? "array" : "scalar";
echo $a[$key];
