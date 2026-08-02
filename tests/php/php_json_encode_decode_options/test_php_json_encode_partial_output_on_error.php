<?php
// vybe-test: php/php_json_encode_decode_options/test_php_json_encode_partial_output_on_error
// origin: languages/php/tests/php/test_php_json_encode_decode_options.rs
// vybe-test-mode: compile

$invalidUtf8 = ["valid" => "text", "invalid" => "\xB1\x31"];
$json = @json_encode($invalidUtf8, JSON_PARTIAL_OUTPUT_ON_ERROR);
echo is_string($json) ? "PARTIAL_JSON" : "FAIL";
