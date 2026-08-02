<?php
// vybe-test: php/json_operations/json_decode_to_stdclass
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$json = '{"title":"Hello","count":5}';
$obj = json_decode($json);
echo $obj->title;
echo $obj->count;
