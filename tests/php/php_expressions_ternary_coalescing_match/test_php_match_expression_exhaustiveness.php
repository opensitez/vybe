<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_match_expression_exhaustiveness
// origin: languages/php/tests/php/test_php_expressions_ternary_coalescing_match.rs
// vybe-test-mode: compile

$state = 1;
$res = match ($state) {
    1 => "One",
    2 => "Two",
    default => "Other",
};
echo $res;
