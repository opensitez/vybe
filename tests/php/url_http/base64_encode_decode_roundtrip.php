<?php
// vybe-test: php/url_http/base64_encode_decode_roundtrip
// origin: languages/php/tests/php/test_url_http.rs
// vybe-test-mode: compile

$original = 'Hello, World! This is a test string.';
$encoded = base64_encode($original);
$decoded = base64_decode($encoded);
echo $decoded === $original ? 'match' : 'mismatch';
echo $encoded;
