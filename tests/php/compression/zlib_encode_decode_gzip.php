<?php
// vybe-test: php/compression/zlib_encode_decode_gzip
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$data = "gzip via zlib_encode";
$encoded = zlib_encode($data, ZLIB_ENCODING_GZIP);
$decoded = zlib_decode($encoded);
echo $decoded === $data ? 'gzip ok' : 'fail';
