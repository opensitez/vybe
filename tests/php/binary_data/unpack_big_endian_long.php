<?php
// vybe-test: php/binary_data/unpack_big_endian_long
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$packed = pack('N', 12345678);
$result = unpack('Nval', $packed);
echo $result['val'];
