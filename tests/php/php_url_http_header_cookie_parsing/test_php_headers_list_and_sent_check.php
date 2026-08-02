<?php
// vybe-test: php/php_url_http_header_cookie_parsing/test_php_headers_list_and_sent_check
// origin: languages/php/tests/php/test_php_url_http_header_cookie_parsing.rs
// vybe-test-mode: compile

echo headers_sent() ? "HEADERS_SENT" : "NOT_SENT";
$headers = headers_list();
echo is_array($headers) ? "ARRAY_OK" : "FAIL";
