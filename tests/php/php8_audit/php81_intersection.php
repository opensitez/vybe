<?php
// vybe-test: php/php8_audit/php81_intersection
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

function foo(A&B $x): C&D { return $x; }
