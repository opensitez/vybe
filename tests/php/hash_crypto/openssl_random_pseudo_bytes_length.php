<?php
// vybe-test: php/hash_crypto/openssl_random_pseudo_bytes_length
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$bytes = openssl_random_pseudo_bytes(16, $strong);
echo strlen($bytes) === 16 ? 'ok' : 'fail';
echo is_bool($strong) ? ':strong flag bool' : ':not bool';
