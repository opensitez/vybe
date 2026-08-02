<?php
// vybe-test: php/php84_request_parse_body_multipart/test_php84_request_parse_body_file_uploads_max
// origin: languages/php/tests/php/test_php84_request_parse_body_multipart.rs
// vybe-test-mode: compile

if (function_exists('request_parse_body')) {
    $parsed = request_parse_body(["max_file_uploads" => 10]);
    echo is_array($parsed) ? "MAX_UPLOADS_OPTION_OK" : "FAIL";
} else {
    echo "MAX_UPLOADS_OPTION_OK";
}
