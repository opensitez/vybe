<?php
// vybe-test: php/php8_audit/php80_named_args
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

function foo($a, $b) {} foo(b: 2, a: 1);
