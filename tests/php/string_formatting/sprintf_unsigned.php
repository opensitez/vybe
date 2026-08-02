<?php
// vybe-test: php/string_formatting/sprintf_unsigned
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo sprintf("%u", 42);
echo "\n";
echo sprintf("%u", PHP_INT_MAX);
echo "\n";
