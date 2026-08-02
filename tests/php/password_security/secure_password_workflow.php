<?php
// vybe-test: php/password_security/secure_password_workflow
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

// Registration
$plaintext = 'user_password_123';
$stored = password_hash($plaintext, PASSWORD_DEFAULT);
// Login verification
$attempt = 'user_password_123';
$valid = password_verify($attempt, $stored);
// Check if needs upgrade
$upgrade = password_needs_rehash($stored, PASSWORD_DEFAULT);
echo $valid ? 'authenticated' : 'denied';
echo !$upgrade ? ':hash current' : ':needs upgrade';
