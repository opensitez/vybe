<?php
// vybe-test: php/binary_data/pack_float
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$packed = pack('f', 3.14);
echo strlen($packed);  // 4 bytes
