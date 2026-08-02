<?php
// vybe-test: php/compression/gzcompress_roundtrip
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$original = str_repeat("compress me! ", 100);
$compressed = gzcompress($original);
$restored = gzuncompress($compressed);
echo $restored === $original ? 'roundtrip ok' : 'fail';
