<?php
// vybe-test: php/php8_audit/php81_enum_method
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

enum Suit: string { case Hearts = 'H'; public function label() { return $this->value; } } echo Suit::Hearts->label();
