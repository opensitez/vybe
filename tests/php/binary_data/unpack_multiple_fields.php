<?php
// vybe-test: php/binary_data/unpack_multiple_fields
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$packed = pack('CCS', 1, 2, 300);
$result = unpack('Cbyte1/Cbyte2/Sshort', $packed);
echo $result['byte1'] . ',' . $result['byte2'] . ',' . $result['short'];
