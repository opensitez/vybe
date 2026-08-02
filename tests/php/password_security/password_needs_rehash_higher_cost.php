<?php
// vybe-test: php/password_security/password_needs_rehash_higher_cost
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$hash = password_hash('pass', PASSWORD_BCRYPT, ['cost' => 10]);
$needs = password_needs_rehash($hash, PASSWORD_BCRYPT, ['cost' => 12]);
echo $needs ? 'needs rehash' : 'ok';
