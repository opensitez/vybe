<?php
// vybe-test: php/php_json_serialization_contracts/test_php_json_last_error_and_msg
// origin: languages/php/tests/php/test_php_json_serialization_contracts.rs
// vybe-test-mode: compile

$result = json_decode("{bad json}");
if (json_last_error() !== JSON_ERROR_NONE) {
    echo "Error code: " . json_last_error() . " Msg: " . json_last_error_msg();
}
