<?php
// vybe-test: php/binary_data/binary_string_substr
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$bin = pack('CCCC', 1, 2, 3, 4);
$part = substr($bin, 1, 2);
$result = unpack('C*', $part);
echo implode(',', $result);
