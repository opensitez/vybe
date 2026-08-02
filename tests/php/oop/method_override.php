<?php
// vybe-test: php/oop/method_override
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class Base { public function greet() { return 'Hello'; } } class Child extends Base { public function greet() { return 'Hi'; } } $c = new Child(); echo $c->greet();
