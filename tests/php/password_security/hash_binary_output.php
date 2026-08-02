<?php
// vybe-test: php/password_security/hash_binary_output
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

$hex = hash('sha256', 'test', false);
$bin = hash('sha256', 'test', true);
echo strlen($hex) === 64 ? 'hex 64 chars' : 'wrong';
echo strlen($bin) === 32 ? ':bin 32 bytes' : ':wrong';
