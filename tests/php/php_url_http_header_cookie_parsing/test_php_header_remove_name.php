<?php
// vybe-test: php/php_url_http_header_cookie_parsing/test_php_header_remove_name
// origin: languages/php/tests/php/test_php_url_http_header_cookie_parsing.rs
// vybe-test-mode: compile

if (!headers_sent()) {
    header("X-Powered-By: Vybe");
    header_remove("X-Powered-By");
}
