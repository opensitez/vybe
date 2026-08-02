<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_null_coalescing_assignment_nested_key
// origin: languages/php/tests/php/test_php_expressions_ternary_coalescing_match.rs
// vybe-test-mode: compile

$settings = [];
$settings["cache"]["ttl"] ??= 3600;
echo $settings["cache"]["ttl"];
