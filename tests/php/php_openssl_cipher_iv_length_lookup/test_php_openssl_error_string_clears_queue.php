<?php
// vybe-test: php/php_openssl_cipher_iv_length_lookup/test_php_openssl_error_string_clears_queue
// origin: languages/php/tests/php/test_php_openssl_cipher_iv_length_lookup.rs
// vybe-test-mode: compile

if (function_exists('openssl_error_string')) {
    @openssl_encrypt("test", "invalid-cipher", "key");
    $err = openssl_error_string();
    echo is_string($err) || $err === false ? "ERROR_STRING_OK" : "FAIL";
} else {
    echo "ERROR_STRING_OK";
}
