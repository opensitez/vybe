<?php
// vybe-test: php/php8_audit/short_closure
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

$fn = fn($x) => $x * 2; echo $fn(5);
