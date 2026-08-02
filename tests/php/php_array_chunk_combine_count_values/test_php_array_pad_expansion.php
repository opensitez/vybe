<?php
// vybe-test: php/php_array_chunk_combine_count_values/test_php_array_pad_expansion
// origin: languages/php/tests/php/test_php_array_chunk_combine_count_values.rs
// vybe-test-mode: compile

$input = [12, 10, 9];
$result = array_pad($input, 5, 0);
echo implode(",", $result);
