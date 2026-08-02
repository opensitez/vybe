<?php
// vybe-test: php/compression/gzcompress_level
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$data = str_repeat("abcdef", 200);
$fast = gzcompress($data, 1);
$best = gzcompress($data, 9);
echo strlen($fast) > 0 ? 'fast ok' : 'fail';
echo strlen($best) > 0 ? ':best ok' : ':fail';
echo strlen($best) <= strlen($fast) ? ':best <= fast' : ':unexpected';
