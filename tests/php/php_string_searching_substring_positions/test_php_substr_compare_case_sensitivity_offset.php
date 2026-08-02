<?php
// vybe-test: php/php_string_searching_substring_positions/test_php_substr_compare_case_sensitivity_offset
// origin: languages/php/tests/php/test_php_string_searching_substring_positions.rs
// vybe-test-mode: compile

$main = "abcde";
$sub = "BC";
echo substr_compare($main, $sub, 1, 2, case_insensitivity: true) === 0 ? "EQUAL" : "NOT_EQUAL";
