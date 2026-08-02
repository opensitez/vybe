<?php
// vybe-test: php/php_reflection_attribute_instantiation_target/test_php_reflection_attribute_on_function
// origin: languages/php/tests/php/test_php_reflection_attribute_instantiation_target.rs
// vybe-test-mode: compile

#[Attribute(Attribute::TARGET_FUNCTION)]
class DeprecatedFunction { public function __construct(public string $reason) {} }

#[DeprecatedFunction("Use newApi() instead")]
function oldApi() {}

$rf = new ReflectionFunction("oldApi");
$attr = $rf->getAttributes(DeprecatedFunction::class)[0];
echo $attr->newInstance()->reason;
