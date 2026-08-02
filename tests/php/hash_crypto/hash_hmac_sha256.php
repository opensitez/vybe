<?php
// vybe-test: php/hash_crypto/hash_hmac_sha256
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$key = 'secret';
$msg = 'data to sign';
$mac = hash_hmac('sha256', $msg, $key);
echo strlen($mac) === 64 ? 'ok' : 'fail';
echo ctype_xdigit($mac) ? ':hex' : ':not hex';
