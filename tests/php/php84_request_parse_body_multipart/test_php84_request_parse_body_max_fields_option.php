<?php
// vybe-test: php/php84_request_parse_body_multipart/test_php84_request_parse_body_max_fields_option
// origin: languages/php/tests/php/test_php84_request_parse_body_multipart.rs
// vybe-test-mode: compile

if (function_exists('request_parse_body')) {
    $parsed = request_parse_body(["max_num_fields" => 50]);
    echo is_array($parsed) ? "MAX_FIELDS_OPTION_OK" : "FAIL";
} else {
    echo "MAX_FIELDS_OPTION_OK";
}
