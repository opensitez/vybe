<?php
// vybe-test: php/url_http/header_send_compile_ok
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

header('Content-Type: application/json');
header('X-Custom-Header: value');
echo 'done';
