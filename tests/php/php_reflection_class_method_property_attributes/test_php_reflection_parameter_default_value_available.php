<?php
// vybe-test: php/php_reflection_class_method_property_attributes/test_php_reflection_parameter_default_value_available
// origin: languages/php/tests/php/test_php_reflection_class_method_property_attributes.rs
// vybe-test-mode: compile

function testParam($a = 100, $b = "default") {}

$rf = new ReflectionFunction("testParam");
$params = $rf->getParameters();
if ($params[0]->isDefaultValueAvailable()) {
    echo $params[0]->getDefaultValue();
}
