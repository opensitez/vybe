<?php
// vybe-test: php/php8_audit/php82_readonly_class
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

readonly class Dto { public function __construct(public string $name) {} }
