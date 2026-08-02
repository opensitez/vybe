<?php
// vybe-test: php/reflection/reflection_class_new_instance_without_constructor
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

class Config { public string $env = 'dev'; }
$rc = new ReflectionClass(Config::class);
$obj = $rc->newInstanceWithoutConstructor();
echo $obj->env;
