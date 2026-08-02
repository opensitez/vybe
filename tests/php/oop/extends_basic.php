<?php
// vybe-test: php/oop/extends_basic
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class Animal { public $name; public function __construct($n) { $this->name = $n; } } class Dog extends Animal { public function speak() { return $this->name . ' barks'; } } $d = new Dog('Rex'); echo $d->speak();
