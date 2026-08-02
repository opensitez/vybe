<?php
// vybe-test: php/php_string_case_folding_mb_case/test_php_string_case_sensitivity_helpers
// origin: languages/php/tests/php/test_php_string_case_folding_mb_case.rs
// vybe-test-mode: compile

$a = "test";
$b = "TEST";
echo strcasecmp($a, $b) === 0 ? "EQUAL_NOCASE" : "NOT_EQUAL";
