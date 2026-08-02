<?php
// vybe-test: php/mb_strings/mb_str_split_chars
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$chars = mb_str_split("hello");
echo implode(',', $chars);
