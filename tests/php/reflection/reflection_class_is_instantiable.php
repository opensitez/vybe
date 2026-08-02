<?php
// vybe-test: php/reflection/reflection_class_is_instantiable
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

abstract class Abs {}
class Concrete {}
$r1 = new ReflectionClass(Abs::class);
$r2 = new ReflectionClass(Concrete::class);
echo $r1->isInstantiable() ? 'yes' : 'no';
echo $r2->isInstantiable() ? 'yes' : 'no';
