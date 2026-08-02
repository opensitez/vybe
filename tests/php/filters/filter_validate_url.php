<?php
// vybe-test: php/filters/filter_validate_url
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

$urls = [
    'https://example.com',
    'http://foo.bar/path?query=1',
    'ftp://files.server.com',
    'not a url',
    '//relative.com',
];
foreach ($urls as $u) {
    echo filter_var($u, FILTER_VALIDATE_URL) !== false ? 'valid' : 'invalid';
    echo ' ';
}
