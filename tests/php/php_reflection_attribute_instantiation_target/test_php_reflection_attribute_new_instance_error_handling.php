<?php
// vybe-test: php/php_reflection_attribute_instantiation_target/test_php_reflection_attribute_new_instance_error_handling
// origin: languages/php/tests/php/test_php_reflection_attribute_instantiation_target.rs
// vybe-test-mode: compile

#[Attribute]
class ConfigAttr { public function __construct(string $required) {} }

#[ConfigAttr] // missing required argument
class Host {}

$rc = new ReflectionClass(Host::class);
$attr = $rc->getAttributes(ConfigAttr::class)[0];
try {
    $attr->newInstance();
} catch (Error $e) {
    echo "Attribute instantiation error caught";
}
