<?php
// vybe-test: php/binary_data/bin2hex_basic
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

echo bin2hex('A');       // 41
echo bin2hex('Hello');   // 48656c6c6f
echo bin2hex("\x00\xFF"); // 00ff
