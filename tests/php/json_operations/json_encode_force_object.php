<?php
// vybe-test: php/json_operations/json_encode_force_object
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$arr = ['apple', 'banana', 'cherry'];
echo json_encode($arr, JSON_FORCE_OBJECT);
