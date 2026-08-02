<?php
// vybe-test: php/binary_data/bin2hex_hex2bin_roundtrip
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$data = "Binary\x00data\xFF\xFE";
echo hex2bin(bin2hex($data)) === $data ? 'roundtrip ok' : 'fail';
