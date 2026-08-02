<?php
// vybe-test: php/binary_data/pack_unsigned_char
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$packed = pack('C', 65);   // 'A'
echo ord($packed) . ':' . $packed;
