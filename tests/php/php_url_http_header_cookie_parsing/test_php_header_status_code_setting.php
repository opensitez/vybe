<?php
// vybe-test: php/php_url_http_header_cookie_parsing/test_php_header_status_code_setting
// origin: languages/php/tests/php/test_php_url_http_header_cookie_parsing.rs
// vybe-test-mode: compile

if (!headers_sent()) {
    header("Content-Type: application/json; charset=UTF-8", replace: true, response_code: 200);
    header("X-Custom-Header: VybeFramework");
}
