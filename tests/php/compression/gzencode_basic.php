<?php
// vybe-test: php/compression/gzencode_basic
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$data = "Hello gzip world!";
$encoded = gzencode($data);
echo strlen($encoded) > 0 ? 'encoded' : 'empty';
// gzip has header magic bytes
echo substr($encoded, 0, 2) === "\x1f\x8b" ? ':gzip magic' : ':no magic';
