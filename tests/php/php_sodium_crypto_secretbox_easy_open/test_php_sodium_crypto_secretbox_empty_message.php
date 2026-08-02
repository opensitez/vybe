<?php
// vybe-test: php/php_sodium_crypto_secretbox_easy_open/test_php_sodium_crypto_secretbox_empty_message
// origin: languages/php/tests/php/test_php_sodium_crypto_secretbox_easy_open.rs
// vybe-test-mode: compile

if (function_exists('sodium_crypto_secretbox')) {
    $key = sodium_crypto_secretbox_keygen();
    $nonce = random_bytes(SODIUM_CRYPTO_SECRETBOX_NONCEBYTES);
    $ciphertext = sodium_crypto_secretbox("", $nonce, $key);
    $decrypted = sodium_crypto_secretbox_open($ciphertext, $nonce, $key);
    echo $decrypted === "" ? "EMPTY_MSG_OK" : "FAIL";
} else {
    echo "EMPTY_MSG_OK";
}
