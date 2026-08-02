<?php
// vybe-test: php/php_mbstring_multibyte_utf8_ops/test_php_mb_substr_count_occurrences
// origin: languages/php/tests/php/test_php_mbstring_multibyte_utf8_ops.rs
// vybe-test-mode: compile

$haystack = "a-b-c-a-b-a";
echo mb_substr_count($haystack, "a", "UTF-8");
