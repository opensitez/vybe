<?php
// vybe-test: php/json_operations/json_encode_booleans
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

echo json_encode(true);
echo json_encode(false);
echo json_encode(['flag' => true, 'other' => false]);
