<?php
// vybe-test: php/binary_data/pack_little_endian_long
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$packed = pack('V', 16909060);  // 0x01020304 little-endian
echo bin2hex($packed);
