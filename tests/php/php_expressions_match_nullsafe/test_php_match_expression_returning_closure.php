<?php
// vybe-test: php/php_expressions_match_nullsafe/test_php_match_expression_returning_closure
// origin: languages/php/tests/php/test_php_expressions_match_nullsafe.rs
// vybe-test-mode: compile

$op = "add";
$handler = match ($op) {
    "add" => fn($a, $b) => $a + $b,
    "sub" => fn($a, $b) => $a - $b,
    default => throw new InvalidArgumentException("Unsupported"),
};
echo $handler(10, 5);
