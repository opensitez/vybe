<?php
// vybe-test: php/binary_data/unpack_unsigned_char
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$packed = pack('C', 42);
$result = unpack('Cval', $packed);
echo $result['val'];
