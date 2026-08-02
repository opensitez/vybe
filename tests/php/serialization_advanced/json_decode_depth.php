<?php
// vybe-test: php/serialization_advanced/json_decode_depth
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$deep = '{"a":{"b":{"c":{"d":1}}}}';
$v = json_decode($deep, true, 512);
echo $v['a']['b']['c']['d'];
