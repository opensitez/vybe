<?php
// vybe-test: php/string_formatting/number_format_custom_separators
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo number_format(1234567.89, 2, ',', '.');  // European format
echo number_format(1234567.89, 2, '.', ' ');  // French format
