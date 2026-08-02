<?php
// vybe-test: php/filters/sanitize_and_validate_pipeline
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

$raw = '  User@Example.COM  ';
$sanitized = strtolower(trim(filter_var($raw, FILTER_SANITIZE_EMAIL)));
$valid = filter_var($sanitized, FILTER_VALIDATE_EMAIL);
echo $valid !== false ? $valid : 'invalid';
