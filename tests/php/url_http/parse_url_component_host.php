<?php
// vybe-test: php/url_http/parse_url_component_host
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$host = parse_url('https://example.com/page', PHP_URL_HOST);
echo $host;
