<?php
// vybe-test: php/binary_data/chr_control_chars
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

echo bin2hex(chr(0));    // 00
echo bin2hex(chr(9));    // 09 (tab)
echo bin2hex(chr(10));   // 0a (newline)
echo bin2hex(chr(13));   // 0d (carriage return)
echo bin2hex(chr(27));   // 1b (escape)
