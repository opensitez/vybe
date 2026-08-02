<?php
// vybe-test: php/compression/gzdeflate_basic
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$data = "deflate compress test";
$deflated = gzdeflate($data);
echo strlen($deflated) > 0 ? 'deflated' : 'empty';
