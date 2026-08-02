<?php
// vybe-test: php/json_operations/json_encode_assoc_array
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$data = ['name' => 'Alice', 'age' => 30, 'active' => true];
echo json_encode($data);
