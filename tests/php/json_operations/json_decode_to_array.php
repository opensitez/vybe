<?php
// vybe-test: php/json_operations/json_decode_to_array
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$json = '{"name":"Alice","age":30}';
$arr = json_decode($json, true);
echo $arr['name'];
echo $arr['age'];
