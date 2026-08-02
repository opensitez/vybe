<?php
// vybe-test: php/string_formatting/chunk_split_base64
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

$data = base64_encode(str_repeat("x", 60));
$formatted = chunk_split($data, 76, "\n");
echo strlen($formatted) > strlen($data) ? 'has newlines' : 'no newlines';
echo "\n";
