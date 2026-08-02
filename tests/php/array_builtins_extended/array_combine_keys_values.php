<?php
// vybe-test: php/array_builtins_extended/array_combine_keys_values
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$keys   = ["a", "b", "c", "d"];
$values = [10, 20, 30, 40];
$map = array_combine($keys, $values);
echo $map["b"];
echo count($map);
