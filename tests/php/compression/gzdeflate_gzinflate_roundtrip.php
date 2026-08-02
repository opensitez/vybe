<?php
// vybe-test: php/compression/gzdeflate_gzinflate_roundtrip
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$original = str_repeat("deflate me ", 50);
$deflated = gzdeflate($original);
$inflated = gzinflate($deflated);
echo $inflated === $original ? 'roundtrip ok' : 'fail';
