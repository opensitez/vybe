<?php
// vybe-test: php/password_security/password_hash_argon2id
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

if (!defined('PASSWORD_ARGON2ID')) {
    echo 'argon2id not available';
} else {
    $hash = password_hash('mypassword', PASSWORD_ARGON2ID);
    echo str_contains($hash, 'argon2id') ? 'argon2id hash' : 'different algo';
}
