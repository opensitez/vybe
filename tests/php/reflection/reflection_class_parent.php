<?php
// vybe-test: php/reflection/reflection_class_parent
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

class Animal {}
class Dog extends Animal {}
$rc = new ReflectionClass(Dog::class);
echo $rc->getParentClass()->getName();
