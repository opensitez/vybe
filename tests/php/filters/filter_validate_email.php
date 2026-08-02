<?php
// vybe-test: php/filters/filter_validate_email
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

$emails = ['user@example.com', 'invalid-email', 'a@b.c', 'missing@', '@nodomain.com'];
foreach ($emails as $e) {
    echo filter_var($e, FILTER_VALIDATE_EMAIL) !== false ? 'valid' : 'invalid';
    echo ' ';
}
