<?php
// vybe-test: php/json_operations/json_encode_numeric_check
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$data = ['count' => '42', 'price' => '9.99'];
echo json_encode($data, JSON_NUMERIC_CHECK);
