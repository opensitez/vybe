<?php
// vybe-test: php/password_security/hash_hmac_tampered
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$key  = 'secret';
$data = 'original';
$sig  = hash_hmac('sha256', $data, $key);
$other = hash_hmac('sha256', 'tampered', $key);
echo hash_equals($sig, $other) ? 'matches' : 'no match';
