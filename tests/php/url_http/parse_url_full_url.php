<?php
// vybe-test: php/url_http/parse_url_full_url
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$parts = parse_url('https://user:pass@example.com:8080/path?q=1#frag');
echo $parts['scheme'];
echo $parts['host'];
echo $parts['port'];
