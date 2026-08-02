<?php
// vybe-test: php/php84_request_parse_body_multipart/test_php84_request_parse_body_null_options_default
// origin: languages/php/tests/php/test_php84_request_parse_body_multipart.rs
// vybe-test-mode: compile

if (function_exists('request_parse_body')) {
    $parsed = request_parse_body(null);
    echo is_array($parsed) ? "NULL_OPTIONS_OK" : "FAIL";
} else {
    echo "NULL_OPTIONS_OK";
}
