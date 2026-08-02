<?php
// vybe-test: php/php_string_case_folding_mb_case/test_php_strtolower_strtoupper_ascii
// origin: languages/php/tests/php/test_php_string_case_folding_mb_case.rs
// vybe-test-mode: compile

$raw = "Laravel Framework 10.x";
echo strtolower($raw) . " " . strtoupper($raw);
