<?php
// vybe-test: php/php_json_encode_decode_options/test_php_json_decode_max_depth_exceeded_error
// origin: languages/php/tests/php/test_php_json_encode_decode_options.rs
// vybe-test-mode: compile

$nestedJson = '{"a":{"b":{"c":{"d":1}}}}';
$res = json_decode($nestedJson, depth: 3);
if ($res === null && json_last_error() === JSON_ERROR_DEPTH) {
    echo "DEPTH_EXCEEDED";
}
