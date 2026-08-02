<?php
// vybe-test: php/compression/compression_empty_string
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$empty = '';
echo gzuncompress(gzcompress($empty)) === $empty ? 'empty compress ok' : 'fail';
echo gzdecode(gzencode($empty))       === $empty ? ':empty gzip ok'    : ':fail';
echo gzinflate(gzdeflate($empty))     === $empty ? ':empty deflate ok' : ':fail';
