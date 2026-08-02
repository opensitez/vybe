<?php
// vybe-test: php/edge_cases/null_coalesce_nested
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

$config = ['db' => ['host' => 'localhost']]; $host = $config['db']['host'] ?? 'default';
