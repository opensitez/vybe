<?php
// vybe-test: php/string_formatting/sprintf_float
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo sprintf("%f", 3.14);
echo "\n";
echo sprintf("%.2f", 3.14159);
echo "\n";
echo sprintf("%.4f", 1.0/3.0);
echo "\n";
