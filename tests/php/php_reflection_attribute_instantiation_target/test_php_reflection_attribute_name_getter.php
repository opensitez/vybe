<?php
// vybe-test: php/php_reflection_attribute_instantiation_target/test_php_reflection_attribute_name_getter
// origin: languages/php/tests/php/test_php_reflection_attribute_instantiation_target.rs
// vybe-test-mode: compile

#[Attribute]
class Component {}

#[Component]
class Service {}

$rc = new ReflectionClass(Service::class);
$attr = $rc->getAttributes()[0];
echo $attr->getName();
