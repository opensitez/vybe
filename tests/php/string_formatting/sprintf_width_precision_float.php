<?php
// vybe-test: php/string_formatting/sprintf_width_precision_float
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo sprintf('%10.2f', 3.14);   // "      3.14"
echo sprintf('%-10.2f|', 3.14); // "3.14      |"
echo sprintf('%010.2f', 3.14);  // "0000003.14"
