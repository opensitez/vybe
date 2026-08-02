<?php
// vybe-test: php/advanced_closures/recursive_closure_via_use_ref
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

$factorial = null;
$factorial = function(int $n) use (&$factorial): int {
    return $n <= 1 ? 1 : $n * $factorial($n - 1);
};
echo $factorial(5);
