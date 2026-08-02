<?php
// vybe-test: php/serialization_advanced/serialize_assoc_array
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$data = ['name' => 'Alice', 'age' => 30, 'active' => true];
$s = serialize($data);
$v = unserialize($s);
echo $v['name'] . ':' . $v['age'];
