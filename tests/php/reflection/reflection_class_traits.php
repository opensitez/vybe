<?php
// vybe-test: php/reflection/reflection_class_traits
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

trait HasLogger { public function log(): void {} }
trait HasCache  { public function cache(): void {} }
class App { use HasLogger, HasCache; }
$rc = new ReflectionClass(App::class);
$traits = array_keys($rc->getTraits());
sort($traits);
echo implode(',', $traits);
