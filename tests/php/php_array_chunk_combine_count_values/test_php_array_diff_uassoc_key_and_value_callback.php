<?php
// vybe-test: php/php_array_chunk_combine_count_values/test_php_array_diff_uassoc_key_and_value_callback
// origin: languages/php/tests/php/test_php_array_chunk_combine_count_values.rs
// vybe-test-mode: compile

$a1 = ["a" => 1, "b" => 2];
$a2 = ["A" => 1, "B" => 2];

$diff = array_diff_uassoc($a1, $a2, fn($k1, $k2) => strcasecmp($k1, $k2));
echo count($diff); // empty because keys match case-insensitively
