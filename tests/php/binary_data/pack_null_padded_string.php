<?php
// vybe-test: php/binary_data/pack_null_padded_string
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$packed = pack('a5', 'hi');
echo strlen($packed) . ':' . bin2hex($packed);
