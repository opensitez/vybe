<?php
// vybe-test: php/php_string_searching_substring_positions/test_php_strripos_case_insensitive_last_position
// origin: languages/php/tests/php/test_php_string_searching_substring_positions.rs
// vybe-test-mode: compile

$haystack = "Abc ABC abc";
$pos = strripos($haystack, "abc");
echo "Pos=$pos";
