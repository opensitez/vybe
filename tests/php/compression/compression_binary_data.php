<?php
// vybe-test: php/compression/compression_binary_data
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$binary = random_bytes(256);
$compressed = gzcompress($binary);
$restored   = gzuncompress($compressed);
echo $restored === $binary ? 'binary roundtrip ok' : 'fail';
