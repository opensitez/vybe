<?php
// vybe-test: php/string_formatting/number_format_small_numbers
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo number_format(0.005, 2);
echo "\n";
echo number_format(0.0,   2);
echo "\n";
echo number_format(-1.5,  1);
echo "\n";
