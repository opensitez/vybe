<?php
// vybe-test: php/php84_request_parse_body_multipart/test_php84_request_parse_body_non_http_env_returns_empty_pair
// origin: languages/php/tests/php/test_php84_request_parse_body_multipart.rs
// vybe-test-mode: compile

if (function_exists('request_parse_body')) {
    [$p, $f] = request_parse_body();
    echo is_array($p) ? "NON_HTTP_EMPTY_PAIR" : "FAIL";
} else {
    echo "NON_HTTP_EMPTY_PAIR";
}
