<?php
// vybe-test: php/edge_cases/null_checks
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

$x = null; if ($x === null) {} if (is_null($x)) {} if (!isset($x)) {}
