<?php
// vybe-test: php/array_extra_builtins/array_key_exists_with_float_key_casting
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = [1.0 => "x", 2 => "y"];
echo array_key_exists(1, $a) ? "one" : "missing";
echo array_key_exists("1", $a) ? "one_str" : "missing_str";
