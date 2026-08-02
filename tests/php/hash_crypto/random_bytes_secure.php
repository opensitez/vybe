<?php
// vybe-test: php/hash_crypto/random_bytes_secure
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$bytes = random_bytes(32);
echo strlen($bytes) === 32 ? 'ok' : 'fail';
echo strlen(bin2hex($bytes)) === 64 ? ':hex ok' : ':hex fail';
