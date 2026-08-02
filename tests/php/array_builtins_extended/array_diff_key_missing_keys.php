<?php
// vybe-test: php/array_builtins_extended/array_diff_key_missing_keys
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = ["x" => 1, "y" => 2, "z" => 3];
$b = ["x" => 99, "z" => 100];
$result = array_diff_key($a, $b);
echo implode(",", array_keys($result));
