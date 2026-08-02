<?php
// vybe-test: php/hash_crypto/hash_raw_output
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$hex = hash('sha256', 'test', false);
$bin = hash('sha256', 'test', true);
echo strlen($hex) === 64 ? 'hex ok' : 'hex fail';
echo strlen($bin) === 32 ? ':bin ok' : ':bin fail';
echo bin2hex($bin) === $hex ? ':round-trip ok' : ':round-trip fail';
