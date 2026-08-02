<?php
// vybe-test: php/php_reflection_attribute_instantiation_target/test_php_reflection_attribute_is_instanceof_check
// origin: languages/php/tests/php/test_php_reflection_attribute_instantiation_target.rs
// vybe-test-mode: compile

interface BaseAttribute {}

#[Attribute]
class CustomAttr implements BaseAttribute {}

#[CustomAttr]
class Target {}

$rc = new ReflectionClass(Target::class);
$attrs = $rc->getAttributes(BaseAttribute::class, ReflectionAttribute::IS_INSTANCEOF);
echo count($attrs);
