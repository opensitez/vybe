<?php
// vybe-test: php/php_filter_validation_sanitization/test_php_filter_var_array_bulk_validation
// origin: languages/php/tests/php/test_php_filter_validation_sanitization.rs
// vybe-test-mode: compile

$data = [
    "email" => "test@domain.org",
    "age" => "25",
    "website" => "http://site.com",
];

$definition = [
    "email" => FILTER_VALIDATE_EMAIL,
    "age" => ["filter" => FILTER_VALIDATE_INT, "options" => ["min_range" => 18]],
    "website" => FILTER_VALIDATE_URL,
];

$validated = filter_var_array($data, $definition);
print_r($validated);
