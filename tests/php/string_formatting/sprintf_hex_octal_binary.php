<?php
// vybe-test: php/string_formatting/sprintf_hex_octal_binary
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

echo sprintf("%x", 255);   // ff
echo sprintf("%X", 255);   // FF
echo sprintf("%o", 8);     // 10
echo sprintf("%b", 10);    // 1010
echo sprintf("%08b", 10);  // 00001010
