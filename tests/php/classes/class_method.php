<?php
// vybe-test: php/classes/class_method
// origin: languages/php/tests/php/test_classes.rs
// vybe-test-mode: compile

class Dog { public $name; public function __construct($n) { $this->name = $n; } public function speak() { return $this->name . ' says Woof'; } } $d = new Dog('Rex'); echo $d->speak();
