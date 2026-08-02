<?php
// vybe-test: php/php_json_encode_decode_options/test_php_json_encode_numeric_keys_preservation
// origin: languages/php/tests/php/test_php_json_encode_decode_options.rs
// vybe-test-mode: compile

$assoc = [1 => "one", 2 => "two"];
echo json_encode($assoc);
