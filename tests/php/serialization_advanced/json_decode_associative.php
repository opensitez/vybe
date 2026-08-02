<?php
// vybe-test: php/serialization_advanced/json_decode_associative
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$json = '{"name":"Bob","scores":[1,2,3]}';
$obj = json_decode($json);
$arr = json_decode($json, true);
echo $obj->name . ':' . implode(',', $arr['scores']);
