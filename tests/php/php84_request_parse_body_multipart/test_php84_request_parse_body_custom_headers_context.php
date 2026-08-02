<?php
// vybe-test: php/php84_request_parse_body_multipart/test_php84_request_parse_body_custom_headers_context
// origin: languages/php/tests/php/test_php84_request_parse_body_multipart.rs
// vybe-test-mode: compile

$_SERVER["HTTP_CONTENT_TYPE"] = "application/x-www-form-urlencoded";
if (function_exists('request_parse_body')) {
    $result = request_parse_body();
    echo is_array($result) ? "CONTENT_TYPE_CONTEXT_OK" : "FAIL";
} else {
    echo "CONTENT_TYPE_CONTEXT_OK";
}
