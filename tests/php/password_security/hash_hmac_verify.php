<?php
// vybe-test: php/password_security/hash_hmac_verify
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$key  = 'my-secret';
$data = 'payload';
$sig  = hash_hmac('sha256', $data, $key);
$expected = hash_hmac('sha256', $data, $key);
echo hash_equals($sig, $expected) ? 'valid' : 'tampered';
