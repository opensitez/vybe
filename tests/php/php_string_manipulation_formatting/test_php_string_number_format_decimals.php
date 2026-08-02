<?php
// vybe-test: php/php_string_manipulation_formatting/test_php_string_number_format_decimals
// origin: languages/php/tests/php/test_php_string_manipulation_formatting.rs
// vybe-test-mode: compile

$number = 1234.5678;
echo number_format($number, 2, '.', ',');
