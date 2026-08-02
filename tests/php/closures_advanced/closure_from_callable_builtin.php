<?php
// vybe-test: php/closures_advanced/closure_from_callable_builtin
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

$upper = Closure::fromCallable('strtoupper');
$len   = Closure::fromCallable('strlen');
echo $upper('hello');
echo $len('hello');
