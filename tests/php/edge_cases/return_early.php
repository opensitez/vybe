<?php
// vybe-test: php/edge_cases/return_early
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

function check($x) { if ($x < 0) return 'negative'; if ($x == 0) return 'zero'; return 'positive'; } echo check(-1);
