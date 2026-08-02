<?php
// vybe-test: php/binary_data/bitwise_toggle_bit
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$flags = 0;
$flags |= (1 << 3);  // set bit 3
echo (bool)($flags & (1 << 3)) ? 'set' : 'clear';
$flags ^= (1 << 3);  // toggle bit 3
echo (bool)($flags & (1 << 3)) ? 'set' : ':clear';
