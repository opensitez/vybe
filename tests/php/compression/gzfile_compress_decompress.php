<?php
// vybe-test: php/compression/gzfile_compress_decompress
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$tmpfile = sys_get_temp_dir() . '/test_' . uniqid() . '.gz';
$data = str_repeat("gzfile test ", 50);
$fh = gzopen($tmpfile, 'w9');
gzwrite($fh, $data);
gzclose($fh);
$fh = gzopen($tmpfile, 'r');
$restored = '';
while (!gzeof($fh)) { $restored .= gzread($fh, 1024); }
gzclose($fh);
@unlink($tmpfile);
echo $restored === $data ? 'file roundtrip ok' : 'fail';
