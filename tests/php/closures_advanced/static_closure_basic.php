<?php
// vybe-test: php/closures_advanced/static_closure_basic
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

$fn = static function(int $n): int { return $n * 2; };
echo $fn(21);
