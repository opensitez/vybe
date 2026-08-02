<?php
// vybe-test: php/reflection/reflection_class_is_interface
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

interface MyInterface {}
class MyClass {}
echo (new ReflectionClass(MyInterface::class))->isInterface() ? 'interface' : 'class';
echo (new ReflectionClass(MyClass::class))->isInterface() ? 'interface' : 'class';
