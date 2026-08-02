<?php
// vybe-test: php/password_security/hash_hmac_basic
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$key = 'secret-key';
$msg = 'message to sign';
$hmac = hash_hmac('sha256', $msg, $key);
echo strlen($hmac) === 64 ? 'hmac length ok' : 'wrong length';
