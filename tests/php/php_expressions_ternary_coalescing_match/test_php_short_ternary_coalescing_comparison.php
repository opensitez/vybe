<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_short_ternary_coalescing_comparison
// origin: languages/php/tests/php/test_php_expressions_ternary_coalescing_match.rs
// vybe-test-mode: compile

$input = "0"; // string "0" is falsy in PHP
$val1 = $input ?: "fallback"; // short ternary triggers fallback
$val2 = $input ?? "fallback"; // null coalescing does NOT trigger (not null)

echo "Ternary=$val1 Coalesce=$val2";
