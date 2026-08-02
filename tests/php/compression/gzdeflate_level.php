<?php
// vybe-test: php/compression/gzdeflate_level
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$data = str_repeat("hello world ", 100);
$d1 = gzdeflate($data, 1);
$d9 = gzdeflate($data, 9);
echo strlen($d9) <= strlen($d1) ? 'level 9 smaller' : 'unexpected';
