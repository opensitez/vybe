<?php
// vybe-test: php/oop/class_constructor
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class Dog { public $name; public function __construct($name) { $this->name = $name; } } $d = new Dog('Rex'); echo $d->name;
