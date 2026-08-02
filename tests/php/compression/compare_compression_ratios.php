<?php
// vybe-test: php/compression/compare_compression_ratios
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$data = str_repeat("The quick brown fox jumps over the lazy dog. ", 100);
$original_size = strlen($data);
$compressed   = strlen(gzcompress($data, 9));
$deflated     = strlen(gzdeflate($data, 9));
$gzipped      = strlen(gzencode($data, 9));
echo $compressed < $original_size ? 'compress ok' : 'fail';
echo $deflated   < $original_size ? ':deflate ok' : ':fail';
echo $gzipped    < $original_size ? ':gzip ok'    : ':fail';
