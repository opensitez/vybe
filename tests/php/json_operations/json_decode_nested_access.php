<?php
// vybe-test: php/json_operations/json_decode_nested_access
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$json = '{"user":{"name":"Carol","scores":[10,20,30]}}';
$data = json_decode($json, true);
echo $data['user']['name'];
echo $data['user']['scores'][1];
