<?php
// vybe-test: php/json_operations/json_decode_depth_limit
// origin: languages/php/tests/php/test_json_operations.rs
// vybe-test-mode: compile

$json = '{"a":{"b":{"c":{"d":"deep"}}}}';
$shallow = json_decode($json, true, 2);
$deep    = json_decode($json, true, 512);
echo is_null($shallow) ? 'null' : 'ok';
echo is_array($deep)   ? 'array' : 'not-array';
