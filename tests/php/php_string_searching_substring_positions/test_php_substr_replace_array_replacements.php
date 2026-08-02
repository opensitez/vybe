<?php
// vybe-test: php/php_string_searching_substring_positions/test_php_substr_replace_array_replacements
// origin: languages/php/tests/php/test_php_string_searching_substring_positions.rs
// vybe-test-mode: compile

$input = ["A: AAA", "B: BBB", "C: CCC"];
$replaced = substr_replace($input, "BBB", 3, 3);
echo implode(",", $replaced);
