<?php
// vybe-test: php/json_operations/json_encode_nested
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$data = [
    'user' => [
        'name' => 'Bob',
        'roles' => ['admin', 'editor'],
        'meta' => ['score' => 99]
    ]
];
echo json_encode($data);
