<?php
// vybe-test: php/scope_patterns/closure_use_by_reference_mutates_outer
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

$total = 0;
$add = function(int $n) use (&$total): void { $total += $n; };
$add(10);
$add(5);
echo $total;
