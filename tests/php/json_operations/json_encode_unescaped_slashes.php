<?php
// vybe-test: php/json_operations/json_encode_unescaped_slashes
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$data = ['url' => 'https://example.com/path/to/resource'];
echo json_encode($data, JSON_UNESCAPED_SLASHES);
