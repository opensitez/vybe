<?php
// vybe-test: php/closures_advanced/closure_from_callable_function
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

function double(int $n): int { return $n * 2; }
$fn = Closure::fromCallable('double');
echo $fn(21);
