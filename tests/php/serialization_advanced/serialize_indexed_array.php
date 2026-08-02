<?php
// vybe-test: php/serialization_advanced/serialize_indexed_array
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$arr = [1, 2, 3, 4, 5];
$s = serialize($arr);
$v = unserialize($s);
echo implode(',', $v);
