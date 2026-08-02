<?php
// vybe-test: php/binary_data/binary_string_search
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$data = pack('CNCN', 0xAA, 0x12345678, 0xBB, 0x87654321);
$pos = strpos($data, chr(0xBB));
echo $pos > 0 ? 'found' : 'not found';
