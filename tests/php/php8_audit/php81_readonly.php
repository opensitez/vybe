<?php
// vybe-test: php/php8_audit/php81_readonly
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

class A { public readonly string $x; public function __construct(string $x) { $this->x = $x; } }
