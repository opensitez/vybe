<?php
// vybe-test: php/compression/compress_json_payload
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$payload = json_encode([
    'users' => array_fill(0, 100, ['name' => 'Alice', 'email' => 'alice@example.com', 'active' => true]),
]);
$compressed = gzencode($payload, 6);
$ratio = strlen($compressed) / strlen($payload);
echo $ratio < 0.5 ? 'good ratio' : 'poor ratio';
echo gzdecode($compressed) === $payload ? ':intact' : ':corrupted';
