<?php
// vybe-test: php/password_security/password_verify_wrong
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$hash = password_hash('rightpassword', PASSWORD_DEFAULT);
echo password_verify('wrongpassword', $hash) ? 'matches' : 'no match';
echo password_verify('',              $hash) ? 'matches' : 'no match';
