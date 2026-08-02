<?php
// vybe-test: php/php_string_searching_substring_positions/test_php_strpbrk_character_set_search
// origin: languages/php/tests/php/test_php_string_searching_substring_positions.rs
// vybe-test-mode: compile

$text = "This is a Simple text.";
echo strpbrk($text, "mi");
