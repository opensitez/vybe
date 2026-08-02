<?php
// vybe-test: php/php8_audit/php80_union_types
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

function foo(int|string $x): int|false { return 0; }
