<?php
// vybe-test: php/string_builtins_extended/number_format_thousands_decimals
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

echo number_format(9876543.21, 2, '.', ',');
echo number_format(1000);
echo number_format(0.125, 3);
