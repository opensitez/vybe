<?php
// vybe-test: php/password_security/password_verify_tampered_hash
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$hash = password_hash('secret', PASSWORD_DEFAULT);
$tampered = substr($hash, 0, 10) . 'XXXX' . substr($hash, 14);
echo password_verify('secret', $tampered) ? 'verified' : 'failed';
