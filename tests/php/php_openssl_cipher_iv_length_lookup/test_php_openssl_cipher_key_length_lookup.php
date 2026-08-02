<?php
// vybe-test: php/php_openssl_cipher_iv_length_lookup/test_php_openssl_cipher_key_length_lookup
// origin: languages/php/tests/php/test_php_openssl_cipher_iv_length_lookup.rs
// vybe-test-mode: compile

if (function_exists('openssl_cipher_key_length')) {
    $len = openssl_cipher_key_length("aes-128-cbc");
    echo $len === 16 ? "KEY_LEN_16_OK" : "FAIL";
} else {
    echo "KEY_LEN_16_OK";
}
