<?php
// vybe-test: php/string_formatting/sprintf_scientific
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo sprintf("%e", 123456.789);  // 1.234568e+5
echo sprintf("%E", 0.000123);    // 1.230000E-4
echo sprintf("%.2e", 1234.5);
echo "\n";
