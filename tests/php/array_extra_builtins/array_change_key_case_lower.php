<?php
// vybe-test: php/array_extra_builtins/array_change_key_case_lower
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = ["FOO" => 10, "Bar" => 20, "baz" => 30];
$lower = array_change_key_case($a, CASE_LOWER);
echo array_key_exists("foo", $lower) ? "yes" : "no";
echo array_key_exists("bar", $lower) ? "yes" : "no";
echo $lower["baz"];
