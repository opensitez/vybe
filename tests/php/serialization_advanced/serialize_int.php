<?php
// vybe-test: php/serialization_advanced/serialize_int
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$s = serialize(42);
$v = unserialize($s);
echo $v;
