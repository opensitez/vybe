<?php
// vybe-test: php/php8_audit/null_coalesce_chain
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

$x = $a ?? $b ?? $c ?? 'last';
