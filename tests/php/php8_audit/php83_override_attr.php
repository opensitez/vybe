<?php
// vybe-test: php/php8_audit/php83_override_attr
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

class B extends A { #[Override] public function foo() {} }
