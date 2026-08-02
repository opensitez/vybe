<?php
// vybe-test: php/serialization_advanced/serialize_bool
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

$t = serialize(true);
$f = serialize(false);
echo unserialize($t) ? 'true' : 'false';
echo unserialize($f) ? 'true' : 'false';
