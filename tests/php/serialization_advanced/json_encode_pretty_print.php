<?php
// vybe-test: php/serialization_advanced/json_encode_pretty_print
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$data = ['name' => 'Alice', 'age' => 30];
$json = json_encode($data, JSON_PRETTY_PRINT);
echo strlen($json) > strlen(json_encode($data)) ? 'pretty' : 'compact';
