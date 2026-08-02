<?php
// vybe-test: php/binary_data/pack_signed_int
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$packed = pack('l', -1);
echo strlen($packed);  // 4 bytes
