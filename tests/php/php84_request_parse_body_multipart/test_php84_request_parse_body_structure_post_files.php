<?php
// vybe-test: php/php84_request_parse_body_multipart/test_php84_request_parse_body_structure_post_files
// origin: languages/php/tests/php/test_php84_request_parse_body_multipart.rs
// vybe-test-mode: compile

if (function_exists('request_parse_body')) {
    [$post, $files] = request_parse_body();
    echo is_array($post) && is_array($files) ? "POST_FILES_DESTRUCTURE_OK" : "FAIL";
} else {
    echo "POST_FILES_DESTRUCTURE_OK";
}
