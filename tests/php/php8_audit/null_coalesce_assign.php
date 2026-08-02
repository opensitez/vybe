<?php
// vybe-test: php/php8_audit/null_coalesce_assign
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

$x = null; $x ??= 'val';
