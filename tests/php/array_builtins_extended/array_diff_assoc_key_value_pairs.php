<?php
// vybe-test: php/array_builtins_extended/array_diff_assoc_key_value_pairs
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = ["color" => "red",  "size" => "M",  "weight" => 10];
$b = ["color" => "red",  "size" => "L",  "weight" => 10];
$diff = array_diff_assoc($a, $b);
echo implode(",", array_keys($diff));
