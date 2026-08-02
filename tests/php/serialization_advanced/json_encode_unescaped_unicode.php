<?php
// vybe-test: php/serialization_advanced/json_encode_unescaped_unicode
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$data = ['greeting' => 'Héllo'];
$escaped = json_encode($data);
$unescaped = json_encode($data, JSON_UNESCAPED_UNICODE);
echo str_contains($unescaped, 'Héllo') ? 'unescaped' : 'escaped';
