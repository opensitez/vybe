<?php
// vybe-test: php/php_url_http_header_cookie_parsing/test_php_parse_url_component_constant_filter
// origin: languages/php/tests/php/test_php_url_http_header_cookie_parsing.rs
// vybe-test-mode: compile

$url = "https://example.com/api/v1";
$host = parse_url($url, PHP_URL_HOST);
$path = parse_url($url, PHP_URL_PATH);
echo "$host $path";
