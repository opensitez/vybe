<?php
// vybe-test: php/password_security/password_verify_correct
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$password = 'correcthorsebatterystaple';
$hash = password_hash($password, PASSWORD_DEFAULT);
echo password_verify($password, $hash) ? 'correct' : 'wrong';
