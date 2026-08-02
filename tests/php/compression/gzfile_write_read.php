<?php
// vybe-test: php/compression/gzfile_write_read
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$tmpfile = sys_get_temp_dir() . '/test_' . uniqid() . '.gz';
$fh = gzopen($tmpfile, 'w');
gzwrite($fh, "line one\n");
gzwrite($fh, "line two\n");
gzclose($fh);
$fh = gzopen($tmpfile, 'r');
$line1 = gzgets($fh);
$line2 = gzgets($fh);
gzclose($fh);
@unlink($tmpfile);
echo trim($line1) . ':' . trim($line2);
