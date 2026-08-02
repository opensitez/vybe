<?php
// vybe-test: php/password_security/password_needs_rehash_different_algo
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$hash = password_hash('pass', PASSWORD_BCRYPT);
$needs = password_needs_rehash($hash, PASSWORD_DEFAULT);
// If DEFAULT changed, might need rehash
echo is_bool($needs) ? 'bool result' : 'not bool';
