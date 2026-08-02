<?php
// vybe-test: php/serialization_advanced/json_encode_unescaped_slashes
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$data = ['url' => 'https://example.com/path'];
$default = json_encode($data);
$noslash = json_encode($data, JSON_UNESCAPED_SLASHES);
echo str_contains($noslash, 'https://example.com/path') ? 'ok' : 'fail';
