<?php
// vybe-test: php/reflection/reflection_has_method_property
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

class Config { public string $env = 'dev'; public function load(): void {} }
$rc = new ReflectionClass(Config::class);
echo $rc->hasMethod('load')      ? 'yes' : 'no';
echo $rc->hasMethod('missing')   ? 'yes' : 'no';
echo $rc->hasProperty('env')     ? 'yes' : 'no';
echo $rc->hasProperty('unknown') ? 'yes' : 'no';
