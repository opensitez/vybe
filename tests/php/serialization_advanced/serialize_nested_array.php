<?php
// vybe-test: php/serialization_advanced/serialize_nested_array
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$nested = ['users' => [['id' => 1, 'name' => 'Alice'], ['id' => 2, 'name' => 'Bob']]];
$s = serialize($nested);
$v = unserialize($s);
echo $v['users'][0]['name'] . ',' . $v['users'][1]['name'];
