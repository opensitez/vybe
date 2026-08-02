<?php
// vybe-test: php/password_security/secure_token_generation
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

function generateToken(int $bytes = 32): string {
    return bin2hex(random_bytes($bytes));
}
$token = generateToken();
echo strlen($token) === 64 ? 'token length ok' : 'wrong length';
echo ctype_xdigit($token) ? ':hex' : ':not hex';
