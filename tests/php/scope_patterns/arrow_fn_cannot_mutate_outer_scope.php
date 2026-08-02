<?php
// vybe-test: php/scope_patterns/arrow_fn_cannot_mutate_outer_scope
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

$x = 10;
$fn = fn() => $x + 1;
$fn();
echo $x;
