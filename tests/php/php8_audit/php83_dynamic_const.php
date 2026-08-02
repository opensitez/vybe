<?php
// vybe-test: php/php8_audit/php83_dynamic_const
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

class A { const X = 1; } $name = 'X'; echo A::X;
