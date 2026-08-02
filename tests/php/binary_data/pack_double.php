<?php
// vybe-test: php/binary_data/pack_double
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$packed = pack('d', 3.14159265358979);
echo strlen($packed);  // 8 bytes
