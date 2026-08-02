<?php
// vybe-test: php/php8_audit/php80_match
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

$x = match(1) { 1 => 'one', 2 => 'two', default => '?' };
