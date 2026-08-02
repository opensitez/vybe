<?php
// vybe-test: php/json_operations/json_encode_unescaped_unicode
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$data = ['greeting' => 'Bonjour'];
echo json_encode($data, JSON_UNESCAPED_UNICODE);
