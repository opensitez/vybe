<?php
// vybe-test: php/php8_audit/php83_closure_from_method
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

class A { public function foo() {} } $fn = (new A())->foo(...);
