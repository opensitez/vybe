<?php
// vybe-test: php/advanced_closures/closure_with_typed_parameter
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

$stringify = function(int|float $n): string { return (string) $n; };
echo $stringify(3.14);
echo $stringify(42);
