<?php
// vybe-test: php/array_builtins_extended/uksort_custom_key_comparator
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = ["cc" => 3, "aaa" => 1, "b" => 2];
uksort($a, fn($x, $y) => strlen($x) - strlen($y));
echo implode(",", array_keys($a));
