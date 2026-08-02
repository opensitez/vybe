<?php
// vybe-test: php/compression/gzcompress_basic
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$data = "Hello, World! This is a test string for compression.";
$compressed = gzcompress($data);
echo strlen($compressed) > 0 ? 'compressed' : 'empty';
echo strlen($compressed) <= strlen($data) ? ':smaller or equal' : ':larger';
