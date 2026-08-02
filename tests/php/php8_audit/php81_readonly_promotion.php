<?php
// vybe-test: php/php8_audit/php81_readonly_promotion
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

class User { public function __construct(public readonly string $name) {} }
