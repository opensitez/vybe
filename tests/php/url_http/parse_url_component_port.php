<?php
// vybe-test: php/url_http/parse_url_component_port
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$port = parse_url('http://example.com:3000/app', PHP_URL_PORT);
echo $port;
