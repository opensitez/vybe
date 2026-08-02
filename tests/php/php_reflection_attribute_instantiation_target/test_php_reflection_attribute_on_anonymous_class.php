<?php
// vybe-test: php/php_reflection_attribute_instantiation_target/test_php_reflection_attribute_on_anonymous_class
// origin: languages/php/tests/php/test_php_reflection_attribute_instantiation_target.rs
// vybe-test-mode: compile

#[Attribute]
class Transient {}

$anon = new #[Transient] class {};
$rc = new ReflectionClass($anon);
echo count($rc->getAttributes(Transient::class));
