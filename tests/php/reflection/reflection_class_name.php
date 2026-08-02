<?php
// vybe-test: php/reflection/reflection_class_name
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

class Foo { public int $x = 1; }
$rc = new ReflectionClass('Foo');
echo $rc->getName();
echo $rc->getShortName();
