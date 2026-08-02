<?php
// vybe-test: php/php8_audit/php80_throw_expr
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

$x = $val ?? throw new Exception('missing');
