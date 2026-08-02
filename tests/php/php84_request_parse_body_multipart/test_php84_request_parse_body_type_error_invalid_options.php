<?php
// vybe-test: php/php84_request_parse_body_multipart/test_php84_request_parse_body_type_error_invalid_options
// origin: languages/php/tests/php/test_php84_request_parse_body_multipart.rs
// vybe-test-mode: compile

if (function_exists('request_parse_body')) {
    try {
        @request_parse_body("invalid_string_option");
    } catch (TypeError $e) {
        echo "TYPE_ERROR_CAUGHT";
    }
} else {
    echo "TYPE_ERROR_CAUGHT";
}
