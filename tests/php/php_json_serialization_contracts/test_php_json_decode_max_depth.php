<?php
// vybe-test: php/php_json_serialization_contracts/test_php_json_decode_max_depth
// origin: languages/php/tests/php/test_php_json_serialization_contracts.rs
// vybe-test-mode: compile

$json = '{"a":{"b":{"c":1}}}';
$obj = json_decode($json, associative: false, depth: 512);
echo $obj->a->b->c;
