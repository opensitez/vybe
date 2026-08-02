<?php
// vybe-test: php/string_interpolation_complex/interpolation_array_in_expression_not_interpolated_without_curly
// origin: languages/php/tests/php/test_string_interpolation_complex.rs
// vybe-test-mode: compile

$arr = [1, 2, 3];
$s = "count is " . count($arr);
echo $s;
echo "\n";
