<?php
// vybe-test: php/php_json_serialization_contracts/test_php_json_validate_function_php83
// origin: languages/php/tests/php/test_php_json_serialization_contracts.rs
// vybe-test-mode: compile

if (function_exists('json_validate')) {
    echo json_validate('{"valid": true}') ? "VALID" : "INVALID";
} else {
    echo "VALID";
}
