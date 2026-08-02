<?php
// vybe-test: php/binary_data/pack_big_endian_long
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$packed = pack('N', 16909060);  // 0x01020304 big-endian
echo bin2hex($packed);
