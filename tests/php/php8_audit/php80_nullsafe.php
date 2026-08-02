<?php
// vybe-test: php/php8_audit/php80_nullsafe
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

$x = $obj?->method()?->prop;
