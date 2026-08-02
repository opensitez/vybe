<?php
// vybe-test: php/compression/gzinflate_basic
// origin: languages/php/tests/php/test_compression.rs
// vybe-test-mode: compile

$data = "inflate this string";
$deflated = gzdeflate($data);
echo gzinflate($deflated);
