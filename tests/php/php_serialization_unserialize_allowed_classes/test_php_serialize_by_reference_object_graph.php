<?php
// vybe-test: php/php_serialization_unserialize_allowed_classes/test_php_serialize_by_reference_object_graph
// origin: languages/php/tests/php/test_php_serialization_unserialize_allowed_classes.rs
// vybe-test-mode: compile

$parent = new stdClass();
$child = new stdClass();
$parent->child = $child;
$child->parent = $parent;

$s = serialize($parent);
$restored = unserialize($s);
echo get_class($restored->child->parent);
