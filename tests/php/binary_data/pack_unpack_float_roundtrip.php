<?php
// vybe-test: php/binary_data/pack_unpack_float_roundtrip
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$val = 3.14;
$packed = pack('d', $val);
$result = unpack('dval', $packed);
echo round($result['val'], 2);
