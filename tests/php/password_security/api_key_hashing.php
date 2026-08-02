<?php
// vybe-test: php/password_security/api_key_hashing
// origin: languages/php/tests/php/test_password_security.rs
// vybe-test-mode: compile

function hashApiKey(string $key): string {
    return hash('sha256', $key);
}
function verifyApiKey(string $key, string $stored): bool {
    return hash_equals($stored, hashApiKey($key));
}
$raw = bin2hex(random_bytes(16));
$stored = hashApiKey($raw);
echo verifyApiKey($raw, $stored)         ? 'valid' : 'invalid';
echo verifyApiKey('wrong-key', $stored)  ? 'valid' : ':invalid';
