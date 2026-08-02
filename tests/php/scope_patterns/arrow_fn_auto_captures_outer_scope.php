<?php
// vybe-test: php/scope_patterns/arrow_fn_auto_captures_outer_scope
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

$multiplier = 7;
$fn = fn(int $x) => $x * $multiplier;
echo $fn(6);
