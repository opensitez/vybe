<?php
// vybe-test: php/url_http/gethostbyname_compile_ok
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$ip = gethostbyname('localhost');
echo is_string($ip) ? 'string' : 'other';
