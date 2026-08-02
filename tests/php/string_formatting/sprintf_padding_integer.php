<?php
// vybe-test: php/string_formatting/sprintf_padding_integer
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo sprintf("%05d", 42);      // 00042
echo sprintf("%-5d|", 42);     // 42   |
echo sprintf("%+d", 42);       // +42
echo sprintf("%+d", -42);      // -42
