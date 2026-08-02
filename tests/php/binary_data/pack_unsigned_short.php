<?php
// vybe-test: php/binary_data/pack_unsigned_short
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$packed = pack('n', 0x0102);  // big-endian unsigned short
echo strlen($packed) . ':' . bin2hex($packed);
