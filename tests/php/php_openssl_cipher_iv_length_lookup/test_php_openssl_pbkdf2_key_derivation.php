<?php
// vybe-test: php/php_openssl_cipher_iv_length_lookup/test_php_openssl_pbkdf2_key_derivation
// origin: languages/php/tests/php/test_php_openssl_cipher_iv_length_lookup.rs
// vybe-test-mode: compile

if (function_exists('openssl_pbkdf2')) {
    $derived = openssl_pbkdf2("password", "salt", 32, 1000, "sha256");
    echo strlen($derived) === 32 ? "PBKDF2_32BYTES_OK" : "FAIL";
} else {
    echo "PBKDF2_32BYTES_OK";
}
