<?php
// vybe-test: php/filters/filter_sanitize_string
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

$raw = '<script>alert("xss")</script>Hello';
$clean = filter_var($raw, FILTER_SANITIZE_SPECIAL_CHARS);
echo strlen($clean) > 0 ? 'sanitized' : 'empty';
