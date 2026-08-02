<?php
// vybe-test: php/serialization_advanced/serialize_string
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$s = serialize("hello world");
$v = unserialize($s);
echo $v;
