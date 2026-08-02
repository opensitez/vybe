<?php
// vybe-test: php/php_reflection_class_method_property_attributes/test_php_reflection_function_invokable
// origin: languages/php/tests/php/test_php_reflection_class_method_property_attributes.rs
// vybe-test-mode: compile

$fn = function(int $a, int $b): int { return $a + $b; };
$rf = new ReflectionFunction($fn);
echo $rf->invoke(10, 20);
