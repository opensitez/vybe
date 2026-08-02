<?php
// vybe-test: php/binary_data/binary_ip_conversion
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$ip = '192.168.1.100';
$parts = explode('.', $ip);
$packed = pack('CCCC', ...$parts);
echo strlen($packed) === 4 ? '4 bytes' : 'wrong';
$unpacked = unpack('C4', $packed);
echo implode('.', $unpacked) === $ip ? ':roundtrip ok' : ':fail';
