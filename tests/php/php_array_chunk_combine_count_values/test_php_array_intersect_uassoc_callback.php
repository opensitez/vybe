<?php
// vybe-test: php/php_array_chunk_combine_count_values/test_php_array_intersect_uassoc_callback
// origin: languages/php/tests/php/test_php_array_chunk_combine_count_values.rs
// vybe-test-mode: compile

$a1 = ["a" => 1, "b" => 2];
$a2 = ["A" => 1, "B" => 3];

$intersect = array_intersect_uassoc($a1, $a2, fn($k1, $k2) => strcasecmp($k1, $k2));
echo count($intersect); // matches "a" => 1
