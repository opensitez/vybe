<?php
// vybe-test: php/php_spl_observer_subject_pattern/test_php_spl_object_storage_serialize_unserialize
// origin: languages/php/tests/php/test_php_spl_observer_subject_pattern.rs
// vybe-test-mode: compile

$s = new SplObjectStorage();
$o = new stdClass(); $o->name = "test";
$s->attach($o, "payload");

$serialized = serialize($s);
$restored = unserialize($serialized);
echo count($restored) === 1 ? "SERIALIZE_OK" : "FAIL";
