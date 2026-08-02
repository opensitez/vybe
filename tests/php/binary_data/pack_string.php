<?php
// vybe-test: php/binary_data/pack_string
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$packed = pack('A5', 'hello');
echo $packed;
$padded = pack('A10', 'hi');
echo strlen($padded);  // 10
