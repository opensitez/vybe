<?php
// vybe-test: php/php_array_chunk_combine_count_values/test_php_array_fill_negative_start_index
// origin: languages/php/tests/php/test_php_array_chunk_combine_count_values.rs
// vybe-test-mode: compile

$a = array_fill(-2, 3, "val");
echo implode(",", array_keys($a)); // -2, 0, 1
