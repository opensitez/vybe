<?php
// vybe-test: php/array_builtins_extended/array_intersect_key_shared_keys
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = ["foo" => 1, "bar" => 2, "baz" => 3];
$b = ["foo" => 99, "baz" => 88];
$result = array_intersect_key($a, $b);
echo implode(",", array_keys($result));
echo implode(",", $result);
