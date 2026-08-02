<?php
// vybe-test: php/compression/zlib_encode_decode_deflate
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$data = str_repeat("zlib test ", 20);
$encoded = zlib_encode($data, ZLIB_ENCODING_DEFLATE);
$decoded = zlib_decode($encoded);
echo $decoded === $data ? 'deflate ok' : 'fail';
