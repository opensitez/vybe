<?php
// vybe-test: php/php_serialization_unserialize_allowed_classes/test_php_serialize_enum_cases
// origin: languages/php/tests/php/test_php_serialization_unserialize_allowed_classes.rs
// vybe-test-mode: compile

enum Role: string { case Admin = "admin"; case User = "user"; }

$s = serialize(Role::Admin);
$restored = unserialize($s);
echo $restored->name . "=" . $restored->value;
