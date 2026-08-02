<?php
// vybe-test: php/serialization_advanced/serialize_null
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$s = serialize(null);
$v = unserialize($s);
var_dump($v);
