<?php
// vybe-test: php/php_password_hashing_cryptography/test_php_hash_pbkdf2_key_derivation
// origin: languages/php/tests/php/test_php_password_hashing_cryptography.rs
// vybe-test-mode: compile

$derived = hash_pbkdf2("sha256", "password", "salt", 1000, 32);
echo strlen($derived) === 64 ? "HEX_LEN_64" : "FAIL";
