<?php
// vybe-test: php/array_builtins_extended/arsort_descending_preserve_keys
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = ["b" => 2, "d" => 4, "a" => 1, "c" => 3];
arsort($a);
echo implode(",", array_keys($a));
echo implode(",", $a);
