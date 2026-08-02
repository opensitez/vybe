<?php
// vybe-test: php/scope_patterns/closure_use_by_value_snapshot
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

$x = 1;
$fn = function() use ($x) { return $x; };
$x = 99;
echo $fn();
