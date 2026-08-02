<?php
// vybe-test: php/classes/class_constructor
// origin: languages/php/tests/php/test_classes.rs
// vybe-test-mode: compile

class Dog { public $name; public function __construct($name) { $this->name = $name; } } $d = new Dog('Rex');
