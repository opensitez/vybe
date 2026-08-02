<?php
// vybe-test: php/php_sodium_crypto_secretbox_easy_open/test_php_sodium_crypto_secretbox_tampered_ciphertext_fails
// origin: languages/php/tests/php/test_php_sodium_crypto_secretbox_easy_open.rs
// vybe-test-mode: compile

if (function_exists('sodium_crypto_secretbox')) {
    $key = sodium_crypto_secretbox_keygen();
    $nonce = random_bytes(SODIUM_CRYPTO_SECRETBOX_NONCEBYTES);
    $ciphertext = sodium_crypto_secretbox("hello", $nonce, $key);
    $ciphertext[0] = chr(ord($ciphertext[0]) ^ 0xff); // Tamper first byte
    $decrypted = sodium_crypto_secretbox_open($ciphertext, $nonce, $key);
    echo $decrypted === false ? "TAMPERED_FAIL_OK" : "FAIL";
} else {
    echo "TAMPERED_FAIL_OK";
}
