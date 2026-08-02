<?php
// vybe-test: php/php_array_chunk_combine_count_values/test_php_array_uintersect_custom_comparator
// origin: languages/php/tests/php/test_php_array_chunk_combine_count_values.rs
// vybe-test-mode: compile

$a1 = ["a" => 1, "b" => 2, "c" => 3];
$a2 = ["x" => 2, "y" => 3, "z" => 4];

$intersect = array_uintersect($a1, $a2, fn($v1, $v2) => $v1 <=> $v2);
echo implode(",", $intersect);
