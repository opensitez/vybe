<?php
// vybe-test: php/php_string_case_folding_mb_case/test_php_mb_case_title_lower_mapping
// origin: languages/php/tests/php/test_php_string_case_folding_mb_case.rs
// vybe-test-mode: compile

$text = "THE QUICK BROWN FOX";
echo mb_convert_case($text, MB_CASE_TITLE, "UTF-8");
