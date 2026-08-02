<?php
// vybe-test: php/php_array_key_operations_indexing/test_php_array_change_key_case_upper_lower
// origin: languages/php/tests/php/test_php_array_key_operations_indexing.rs
// vybe-test-mode: compile

$input = ["First" => 1, "SecOND" => 4];
$lower = array_change_key_case($input, CASE_LOWER);
$upper = array_change_key_case($input, CASE_UPPER);
echo implode(",", array_keys($lower)) . " | " . implode(",", array_keys($upper));
