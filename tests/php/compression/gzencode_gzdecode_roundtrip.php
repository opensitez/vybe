<?php
// vybe-test: php/compression/gzencode_gzdecode_roundtrip
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$original = str_repeat("gzip test data ", 50);
$encoded = gzencode($original);
$decoded = gzdecode($encoded);
echo $decoded === $original ? 'roundtrip ok' : 'fail';
