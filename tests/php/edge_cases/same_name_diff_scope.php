<?php
// vybe-test: php/edge_cases/same_name_diff_scope
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

$x = 'global'; function foo() { $x = 'local'; return $x; } echo foo(); echo $x;
