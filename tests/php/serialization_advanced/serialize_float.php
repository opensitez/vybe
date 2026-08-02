<?php
// vybe-test: php/serialization_advanced/serialize_float
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$s = serialize(3.14);
$v = unserialize($s);
echo $v;
