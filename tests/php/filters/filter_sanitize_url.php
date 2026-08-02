<?php
// vybe-test: php/filters/filter_sanitize_url
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

$url = "https://example.com/path with spaces?q=hello world";
$clean = filter_var($url, FILTER_SANITIZE_URL);
echo str_contains($clean, 'example.com') ? 'ok' : 'fail';
