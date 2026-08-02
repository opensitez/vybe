<?php
// vybe-test: php/url_http/headers_sent_compile_ok
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$sent = headers_sent();
echo is_bool($sent) ? 'bool' : 'other';
