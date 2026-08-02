<?php
// vybe-test: php/php_reflection_attribute_instantiation_target/test_php_reflection_attribute_on_function_parameter
// origin: languages/php/tests/php/test_php_reflection_attribute_instantiation_target.rs
// vybe-test-mode: compile

#[Attribute(Attribute::TARGET_PARAMETER)]
class ValidateEmail {}

function registerUser(#[ValidateEmail] string $email) {}

$rp = new ReflectionParameter("registerUser", "email");
$attrs = $rp->getAttributes(ValidateEmail::class);
echo count($attrs);
