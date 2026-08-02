<?php
// vybe-test: php/edge_cases/chained_coalesce
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

$x = $a ?? $b ?? $c ?? 'default';
