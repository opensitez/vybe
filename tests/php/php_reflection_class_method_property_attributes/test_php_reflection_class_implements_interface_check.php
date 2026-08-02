<?php
// vybe-test: php/php_reflection_class_method_property_attributes/test_php_reflection_class_implements_interface_check
// origin: languages/php/tests/php/test_php_reflection_class_method_property_attributes.rs
// vybe-test-mode: compile

interface PluginInterface {}
class MyPlugin implements PluginInterface {}

$rc = new ReflectionClass(MyPlugin::class);
echo $rc->implementsInterface(PluginInterface::class) ? "IMPLEMENTS" : "NO";
