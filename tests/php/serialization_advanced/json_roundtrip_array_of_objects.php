<?php
// vybe-test: php/serialization_advanced/json_roundtrip_array_of_objects
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$users = [
    ['id' => 1, 'name' => 'Alice', 'active' => true],
    ['id' => 2, 'name' => 'Bob',   'active' => false],
];
$json = json_encode($users);
$decoded = json_decode($json, true);
echo $decoded[0]['name'] . ',' . $decoded[1]['name'];
