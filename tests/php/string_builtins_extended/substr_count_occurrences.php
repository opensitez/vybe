<?php
// vybe-test: php/string_builtins_extended/substr_count_occurrences
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

echo substr_count("hello world hello world", "hello");
echo substr_count("banana", "ana");
echo substr_count("aababab", "ab");
