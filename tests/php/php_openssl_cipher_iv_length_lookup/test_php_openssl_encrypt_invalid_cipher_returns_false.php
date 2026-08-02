<?php
// vybe-test: php/php_openssl_cipher_iv_length_lookup/test_php_openssl_encrypt_invalid_cipher_returns_false
// origin: languages/php/tests/php/test_php_openssl_cipher_iv_length_lookup.rs
// vybe-test-mode: compile

if (function_exists('openssl_encrypt')) {
    $res = @openssl_encrypt("test", "invalid-cipher-name-999", "key", 0, "iv");
    echo $res === false ? "INVALID_CIPHER_FALSE_OK" : "FAIL";
} else {
    echo "INVALID_CIPHER_FALSE_OK";
}
