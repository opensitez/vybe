<?php
// vybe-test: php/json_operations/json_encode_pretty_print
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$data = ['key' => 'value', 'num' => 42];
echo json_encode($data, JSON_PRETTY_PRINT);
