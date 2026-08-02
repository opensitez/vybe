<?php
// vybe-test: php/array_builtins_extended/array_diff_missing_values
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = ["apple", "banana", "cherry", "date"];
$b = ["banana", "date", "elderberry"];
$diff = array_diff($a, $b);
echo implode(",", $diff);
