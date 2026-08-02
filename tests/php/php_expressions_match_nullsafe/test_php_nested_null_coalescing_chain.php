<?php
// vybe-test: php/php_expressions_match_nullsafe/test_php_nested_null_coalescing_chain
// origin: languages/php/tests/php/test_php_expressions_match_nullsafe.rs
// vybe-test-mode: compile

$opt1 = null;
$opt2 = null;
$opt3 = "final_fallback";
$res = $opt1 ?? $opt2 ?? $opt3;
echo $res;
