<?php
// vybe-test: php/php_array_key_operations_indexing/test_php_array_fill_zero_based_index
// origin: languages/php/tests/php/test_php_array_key_operations_indexing.rs
// vybe-test-mode: compile

$a = array_fill(5, 3, "banana");
echo implode(",", array_keys($a));
