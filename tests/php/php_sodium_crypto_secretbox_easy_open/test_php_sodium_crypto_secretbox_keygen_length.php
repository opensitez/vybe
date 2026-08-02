<?php
// vybe-test: php/php_sodium_crypto_secretbox_easy_open/test_php_sodium_crypto_secretbox_keygen_length
// origin: languages/php/tests/php/test_php_sodium_crypto_secretbox_easy_open.rs
// vybe-test-mode: compile

if (function_exists('sodium_crypto_secretbox_keygen')) {
    $key = sodium_crypto_secretbox_keygen();
    echo strlen($key) === SODIUM_CRYPTO_SECRETBOX_KEYBYTES ? "SECRETBOX_KEYBYTES_OK" : "FAIL";
} else {
    echo "SECRETBOX_KEYBYTES_OK";
}
