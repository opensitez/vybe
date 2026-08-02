<?php
// vybe-test: php/hash_crypto/base64_roundtrip_binary
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$raw = random_bytes(24);
$encoded = base64_encode($raw);
$decoded = base64_decode($encoded);
echo $decoded === $raw ? 'roundtrip ok' : 'roundtrip fail';
echo ctype_print($encoded) ? ':printable' : ':not printable';
