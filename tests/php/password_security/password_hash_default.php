<?php
// vybe-test: php/password_security/password_hash_default
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$hash = password_hash('secret123', PASSWORD_DEFAULT);
echo is_string($hash) ? 'hashed' : 'fail';
echo strlen($hash) > 20 ? ':long enough' : ':too short';
