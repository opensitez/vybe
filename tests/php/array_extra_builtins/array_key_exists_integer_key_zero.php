<?php
// vybe-test: php/array_extra_builtins/array_key_exists_integer_key_zero
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = [0 => "first", 1 => "second", 2 => "third"];
echo array_key_exists(0, $a) ? "exists" : "missing";
echo array_key_exists(3, $a) ? "exists" : "missing";
$empty = [];
echo array_key_exists(0, $empty) ? "exists" : "missing";
