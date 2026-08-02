<?php
// vybe-test: php/string_formatting/sprintf_width_dynamic
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo sprintf('%*d', 5, 42);   // PHP uses %5d style; *-width is non-standard but %5d works
echo sprintf('%5d', 42);
echo "\n";
