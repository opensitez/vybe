<?php
// vybe-test: php/oop/parent_call
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class Base { public function foo() { return 'base'; } } class Child extends Base { public function foo() { return parent::foo() . '+child'; } } $c = new Child();
