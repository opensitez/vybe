<?php
// vybe-test: php/mb_strings/mb_substr_count_basic
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

echo mb_substr_count("hello world hello", "hello");
echo mb_substr_count("abababab", "ab");
