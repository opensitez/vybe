<?php
// vybe-test: php/binary_data/unpack_array_format
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$packed = pack('C4', 10, 20, 30, 40);
$result = unpack('C4bytes', $packed);
echo implode(',', $result);
