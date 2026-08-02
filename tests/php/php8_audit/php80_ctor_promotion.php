<?php
// vybe-test: php/php8_audit/php80_ctor_promotion
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

class P { public function __construct(public int $x, public string $y = 'hi') {} } new P(1);
