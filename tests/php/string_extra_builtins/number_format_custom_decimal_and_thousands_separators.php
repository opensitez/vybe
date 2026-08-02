<?php
// vybe-test: php/string_extra_builtins/number_format_custom_decimal_and_thousands_separators
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$n = 1234567.891;
echo number_format($n, 2, ',', '.');
echo number_format($n, 3, '/', '_');
echo number_format(0.5, 1, ',', '');
