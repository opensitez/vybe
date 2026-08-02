<?php
// vybe-test: php/php_url_http_header_cookie_parsing/test_php_http_response_code_get_set
// origin: languages/php/tests/php/test_php_url_http_header_cookie_parsing.rs
// vybe-test-mode: compile

if (!headers_sent()) {
    http_response_code(404);
    echo http_response_code();
}
