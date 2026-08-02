<?php
// vybe-test: php/json_operations/json_encode_empty_array_vs_object
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$arr = [];
$obj = new stdClass();
echo json_encode($arr);
echo json_encode($obj);
