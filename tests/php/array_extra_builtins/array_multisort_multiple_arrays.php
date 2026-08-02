<?php
// vybe-test: php/array_extra_builtins/array_multisort_multiple_arrays
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$data = [3, 1, 4, 1, 5, 9, 2, 6];
$keys = ["c", "a", "d", "a2", "e", "i", "b", "f"];
array_multisort($data, SORT_ASC, $keys);
echo count($data);
echo is_array($keys) ? "array" : "not";
