<?php
// vybe-test: php/password_security/password_hash_argon2i
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

if (!defined('PASSWORD_ARGON2I')) {
    echo 'argon2i not available';
} else {
    $hash = password_hash('mypassword', PASSWORD_ARGON2I);
    echo str_contains($hash, 'argon2i') ? 'argon2i hash' : 'different algo';
}
