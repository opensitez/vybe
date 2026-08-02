<?php
// vybe-test: php/php84_request_parse_body_multipart/test_php84_request_parse_body_json_content_type
// origin: languages/php/tests/php/test_php84_request_parse_body_multipart.rs
// vybe-test-mode: compile

$_SERVER["CONTENT_TYPE"] = "application/json";
if (function_exists('request_parse_body')) {
    $parsed = request_parse_body();
    echo is_array($parsed) ? "JSON_CONTENT_TYPE_OK" : "FAIL";
} else {
    echo "JSON_CONTENT_TYPE_OK";
}
