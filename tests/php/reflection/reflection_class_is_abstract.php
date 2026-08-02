<?php
// vybe-test: php/reflection/reflection_class_is_abstract
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

abstract class Base {}
class Concrete extends Base {}
$rb = new ReflectionClass(Base::class);
$rc = new ReflectionClass(Concrete::class);
echo $rb->isAbstract() ? 'abstract' : 'concrete';
echo $rc->isAbstract() ? 'abstract' : 'concrete';
