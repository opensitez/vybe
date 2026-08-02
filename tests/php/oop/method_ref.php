<?php
// vybe-test: php/oop/method_ref
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class A { public function foo() { return 42; } } $a = new A(); $fn = $a->foo(...);
