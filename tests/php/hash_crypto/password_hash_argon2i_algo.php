<?php
// vybe-test: php/hash_crypto/password_hash_argon2i_algo
// origin: languages/php/tests/php/test_hash_crypto.rs
// vybe-test-mode: compile

if (defined('PASSWORD_ARGON2I')) {
    $hash = password_hash('mypassword', PASSWORD_ARGON2I);
    echo is_string($hash) ? 'hashed' : 'fail';
} else {
    echo 'argon2i unavailable';
}
