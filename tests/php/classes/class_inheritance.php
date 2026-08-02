<?php
// vybe-test: php/classes/class_inheritance
// origin: languages/php/tests/php/test_classes.rs
// vybe-test-mode: compile

class Animal { public $name; public function __construct($n) { $this->name = $n; } } class Cat extends Animal { public function speak() { return $this->name . ' says Meow'; } } $c = new Cat('Whiskers'); echo $c->speak();
