<?php
// vybe-test: php/php_serialization_unserialize_allowed_classes/test_php_serialize_stdclass_object
// origin: languages/php/tests/php/test_php_serialization_unserialize_allowed_classes.rs
// vybe-test-mode: compile

$obj = new stdClass();
$obj->title = "Test";
$obj->tags = ["php", "unit"];

$serialized = serialize($obj);
$restored = unserialize($serialized);
echo $restored->title . " tags=" . count($restored->tags);
