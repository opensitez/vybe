<?php
// vybe-test: php/php8_audit/php80_trailing_comma_use
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

$fn = function() use ($a, $b,) {};
