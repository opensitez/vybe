<?php
// vybe-test: php/php_string_case_folding_mb_case/test_php_mb_convert_case_upper_lower_modes
// origin: languages/php/tests/php/test_php_string_case_folding_mb_case.rs
// vybe-test-mode: compile

$text = "éclair";
echo mb_convert_case($text, MB_CASE_UPPER, "UTF-8") . " " . mb_convert_case("ÉCLAIR", MB_CASE_LOWER, "UTF-8");
