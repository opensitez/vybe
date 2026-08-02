<?php
// vybe-test: php/array_extra_builtins/shuffle_randomize_array_order
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = [1, 2, 3, 4, 5, 6];
shuffle($a);
echo count($a);
echo is_array($a) ? "array" : "not";
