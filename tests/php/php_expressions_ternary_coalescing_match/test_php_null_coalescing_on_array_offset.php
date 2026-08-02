<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_null_coalescing_on_array_offset
// origin: languages/php/tests/php/test_php_expressions_ternary_coalescing_match.rs
// vybe-test-mode: compile

$arr = ["a" => 10];
echo $arr["a"] ?? 0;
echo $arr["b"] ?? 0;
