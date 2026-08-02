<?php
// vybe-test: php/type_juggling/intval_bases
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

echo intval('0x1A', 16) . "\n";  // 26
echo intval('0b1010', 2) . "\n"; // 10
echo intval('077', 8) . "\n";    // 63
echo intval('42', 10) . "\n";    // 42
echo intval('42abc') . "\n";     // 42
