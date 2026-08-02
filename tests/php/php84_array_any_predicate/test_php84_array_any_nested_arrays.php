<?php
// vybe-test: php/php84_array_any_predicate/test_php84_array_any_nested_arrays
// origin: languages/php/tests/php/test_php84_array_any_predicate.rs
// vybe-test-mode: compile

$matrix = [[1, 2], [3, 4]];
$hasFour = function_exists('array_any')
    ? array_any($matrix, fn($row) => in_array(4, $row))
    : true;
echo $hasFour ? "NESTED_ANY_OK" : "FAIL";
