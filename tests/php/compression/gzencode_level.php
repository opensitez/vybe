<?php
// vybe-test: php/compression/gzencode_level
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$data = str_repeat("level test ", 100);
$low  = gzencode($data, 1);
$high = gzencode($data, 9);
echo strlen($low)  > 0 ? 'low ok'  : 'fail';
echo strlen($high) > 0 ? ':high ok' : ':fail';
