<?php
// vybe-test: php/url_http/http_response_code_compile_ok
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$code = http_response_code();
echo is_int($code) || $code === false ? 'ok' : 'fail';
