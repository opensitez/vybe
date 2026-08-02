<?php
// vybe-test: php/php_json_encode_decode_options/test_php_json_exception_code_and_message
// origin: languages/php/tests/php/test_php_json_encode_decode_options.rs
// vybe-test-mode: compile

try {
    json_decode("{malformed}", flags: JSON_THROW_ON_ERROR);
} catch (JsonException $e) {
    echo "Code=" . $e->getCode() . " Msg=" . $e->getMessage();
}
