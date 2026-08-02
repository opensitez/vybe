<?php
// vybe-test: php/hash_crypto/hash_sha256_algo
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$h = hash('sha256', 'hello world');
echo strlen($h) === 64 ? 'ok' : 'fail';
echo ctype_xdigit($h) ? ':hex' : ':not hex';
