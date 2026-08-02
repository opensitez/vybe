<?php
// vybe-test: php/reflection/reflection_property_get_set
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

class Box { public int $width = 10; public int $height = 20; }
$obj = new Box();
$rp = new ReflectionProperty(Box::class, 'width');
echo $rp->getValue($obj);
$rp->setValue($obj, 50);
echo $obj->width;
