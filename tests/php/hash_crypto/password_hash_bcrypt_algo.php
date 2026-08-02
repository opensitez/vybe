<?php
// vybe-test: php/hash_crypto/password_hash_bcrypt_algo
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

$hash = password_hash('mypassword', PASSWORD_BCRYPT);
echo is_string($hash) ? 'hashed' : 'fail';
echo str_starts_with($hash, '$2y$') ? ':bcrypt' : ':other';
