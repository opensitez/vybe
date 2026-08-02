<?php
// vybe-test: php/reflection/reflection_class_constants
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

class Status {
    const OK      = 200;
    const CREATED = 201;
    const ERROR   = 500;
}
$rc = new ReflectionClass(Status::class);
$consts = $rc->getConstants();
echo count($consts);
echo $consts['OK'];
