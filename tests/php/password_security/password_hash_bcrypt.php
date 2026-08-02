<?php
// vybe-test: php/password_security/password_hash_bcrypt
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$hash = password_hash('mypassword', PASSWORD_BCRYPT);
echo str_starts_with($hash, '$2y$') ? 'bcrypt hash' : 'not bcrypt';
