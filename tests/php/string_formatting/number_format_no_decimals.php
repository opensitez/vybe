<?php
// vybe-test: php/string_formatting/number_format_no_decimals
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo number_format(9999999, 0, '.', ',');
echo "\n";
