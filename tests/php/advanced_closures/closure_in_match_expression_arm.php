<?php
// vybe-test: php/advanced_closures/closure_in_match_expression_arm
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

$op = 'double';
$fn = match ($op) {
    'double' => fn(int $x) => $x * 2,
    'square' => fn(int $x) => $x * $x,
    default  => fn(int $x) => $x,
};
echo $fn(7);
