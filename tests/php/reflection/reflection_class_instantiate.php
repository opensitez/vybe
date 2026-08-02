<?php
// vybe-test: php/reflection/reflection_class_instantiate
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

class Point { public function __construct(public int $x, public int $y) {} }
$rc = new ReflectionClass(Point::class);
$obj = $rc->newInstance(3, 7);
echo $obj->x . ',' . $obj->y;
