<?php
// vybe-test: php/compression/gzdecode_basic
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$data = "decode this";
$encoded = gzencode($data);
$decoded = gzdecode($encoded);
echo $decoded;
