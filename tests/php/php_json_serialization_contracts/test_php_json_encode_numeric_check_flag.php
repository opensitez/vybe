<?php
// vybe-test: php/php_json_serialization_contracts/test_php_json_encode_numeric_check_flag
// origin: languages/php/tests/php/test_php_json_serialization_contracts.rs
// vybe-test-mode: compile

$data = ["id" => "123", "score" => "98.6", "name" => "Alice"];
echo json_encode($data, JSON_NUMERIC_CHECK);
