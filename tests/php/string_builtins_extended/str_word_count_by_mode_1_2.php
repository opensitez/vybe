<?php
// vybe-test: php/string_builtins_extended/str_word_count_by_mode_1_2
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

echo implode("|", str_word_count("one two-three", 1));
echo "|";
echo implode("|", array_keys(str_word_count("one two-three", 2)));
